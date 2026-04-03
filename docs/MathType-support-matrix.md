# MathType support matrix

This file is generated from the current codebase and then manually curated as implementation progresses.

## Snapshot
- Declared official TM_* templates in `MtefRecord`: **39**
- Templates with both builder + writer path: **38**
- Templates with builder helper only: **0**
- Templates only declared in constants: **1**

## Parser coverage signals
- linear commands: \frac, \sqrt, \sum, \int, \xrightarrow, \xleftarrow, \overbrace, \overarc
- environments: matrix, pmatrix, bmatrix, cases, aligned, align / align*, split, longdivision

## Official MTEF template coverage
| Template | Selector | Status | Builder path | Writer path |
|---|---:|---|---|---|
| `TM_ANGLE` | 0 | implemented | `writeAngleHeader` | `writeParenFence / writeFenceTemplate` |
| `TM_PAREN` | 1 | implemented | `writeParenHeader` | `writeParenFence / writeCommandNode` |
| `TM_BRACE` | 2 | implemented | `writeBraceHeader` | `writeCommandNode / writeCasesAsBraceFence` |
| `TM_BRACK` | 3 | implemented | `writeBracketHeader` | `writeParenFence / writeCommandNode` |
| `TM_BAR` | 4 | implemented | `writeBarHeader` | `writeCommandNode / writeFenceTemplate` |
| `TM_DBAR` | 5 | implemented | `writeDoubleBarHeader` | `writeFenceTemplate` |
| `TM_FLOOR` | 6 | implemented | `writeFloorHeader` | `writeFenceTemplate` |
| `TM_CEILING` | 7 | implemented | `writeCeilingHeader` | `writeFenceTemplate` |
| `TM_OBRACK` | 8 | implemented | `writeOpenBracketHeader` | `writeFenceTemplate` |
| `TM_INTERVAL` | 9 | implemented | `writeIntervalHeader` | `writeFenceTemplate` |
| `TM_ROOT` | 10 | implemented | `writeSqrtHeader/writeNthRootHeader` | `writeSqrtNode` |
| `TM_FRACT` | 11 | implemented | `writeFractionHeader` | `writeFractionNode` |
| `TM_UBAR` | 12 | implemented | `writeUnderlineHeader` | `writeCommandNode` |
| `TM_OBAR` | 13 | implemented | `writeOverlineHeader` | `writeCommandNode` |
| `TM_ARROW` | 14 | implemented | `writeArrowHeader` | `writeArrowNode` |
| `TM_INTEG` | 15 | implemented | `writeIntegralHeader` | `writeBigOpHeader (integrals)` |
| `TM_SUM` | 16 | implemented | `writeSumHeader` | `writeBigOpHeader (sum)` |
| `TM_PROD` | 17 | implemented | `writeProductHeader` | `writeBigOpHeader (prod)` |
| `TM_COPROD` | 18 | implemented | `writeCoproductHeader` | `writeBigOpHeader (coprod)` |
| `TM_UNION` | 19 | implemented | `writeUnionHeader` | `writeBigOpHeader (bigcup)` |
| `TM_INTER` | 20 | implemented | `writeIntersectionHeader` | `writeBigOpHeader (bigcap)` |
| `TM_INTOP` | 21 | implemented | `writeIntegralStyleBigOpHeader` | `writeBigOpHeader (intop)` |
| `TM_SUMOP` | 22 | implemented | `writeSummationStyleBigOpHeader` | `writeBigOpHeader (sumop-like)` |
| `TM_LIM` | 23 | implemented | `writeLimitHeader` | `writeLimitComplete` |
| `TM_HBRACE` | 24 | implemented | `writeHBraceHeader` | `writeHorizontalBrace` |
| `TM_HBRACK` | 25 | implemented | `writeHBrackHeader` | `writeHorizontalBracket` |
| `TM_LDIV` | 26 | implemented | `writeLongDivisionHeader` | `writeLongDivisionNode` |
| `TM_SUB` | 27 | implemented | `writeSubscriptHeader` | `writeSubscriptNode / writeLeadingScriptAttachment` |
| `TM_SUP` | 28 | implemented | `writeSuperscriptHeader` | `writeSuperscriptNode / writeLeadingScriptAttachment` |
| `TM_SUBSUP` | 29 | implemented | `writeSubSuperscriptHeader` | `writeSupSubAttachment / writeLeadingScriptAttachment` |
| `TM_DIRAC` | 30 | implemented | `writeDiracHeader` | `writeDiracNode` |
| `TM_VEC` | 31 | implemented | `writeVecHeader` | `writeCommandNode` |
| `TM_TILDE` | 32 | implemented | `writeTildeHeader` | `writeCommandNode` |
| `TM_HAT` | 33 | implemented | `writeHatHeader` | `writeCommandNode` |
| `TM_ARC` | 34 | implemented | `writeArcHeader` | `writeArcNode / writeCommandNode` |
| `TM_JSTATUS` | 35 | implemented | `writeJointStatusHeader` | `writeCommandNode` |
| `TM_STRIKE` | 36 | implemented | `writeStrikeHeader` | `writeStrikeNode / writeCommandNode` |
| `TM_BOX` | 37 | implemented | `writeBoxHeader` | `writeBoxNode / writeCommandNode` |
| `TM_LSCRIPT` | 44 | declared-only | `—` | `—` |

## Notes
- `implemented` means the repository contains both a template builder helper and an explicit writer path using that template family.
- `builder-only` means helper code exists but no explicit writer use was mapped in this first pass.
- `declared-only` means the official template constant exists in `MtefRecord`, but no builder/writer path was found yet.
- Left-script / prescript input such as `{}^{a}x`, `{}_{b}x`, `{}_{b}^{a}x` is emitted in **MTEF v5 form** via `TM_SUB` / `TM_SUP` / `TM_SUBSUP` plus `TV_SU_PRECEDES`, rather than the legacy `TM_LSCRIPT(44)` selector.
- `\xrightarrow` / `\xleftarrow` now follow the official amsmath signature `\xrightarrow[below]{above}` / `\xleftarrow[below]{above}` for the supported single-arrow family.
- `TM_ARC` currently covers the documented over-arc family (`\arc`, `\overarc`, `\overparen`, `\wideparen`). `\underarc` is intentionally not claimed here because it is not part of amsmath and no official MathType mapping has been confirmed in this repository yet.
- This matrix measures **code-path existence**, not semantic completeness. For example, integrals may be implemented while still missing full variation coverage (double/triple/loop), and some behaviors still rely on Linux-side byte-level validation instead of a fresh official Windows round-trip.
