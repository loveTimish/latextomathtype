# MathType support matrix

This file is generated from the current codebase and then manually curated as implementation progresses.

## Snapshot
- Declared official TM_* templates in `MtefRecord`: **39**
- Templates with both builder + writer path: **18**
- Templates with builder helper only: **0**
- Templates only declared in constants: **21**

## Parser coverage signals
- linear: no direct hits
- environments: matrix, pmatrix, bmatrix, cases, aligned, align, align*, split, longdivision

## Official MTEF template coverage
| Template | Selector | Status | Builder path | Writer path |
|---|---:|---|---|---|
| `TM_ANGLE` | 0 | declared-only | `—` | `—` |
| `TM_PAREN` | 1 | implemented | `writeParenHeader` | `writeParenFence / writeCommandNode` |
| `TM_BRACE` | 2 | implemented | `writeBraceHeader` | `writeCommandNode` |
| `TM_BRACK` | 3 | implemented | `writeBracketHeader` | `writeParenFence / writeCommandNode` |
| `TM_BAR` | 4 | implemented | `writeBarHeader` | `writeCommandNode` |
| `TM_DBAR` | 5 | implemented | `—` | `fence header writer (single-sided)` |
| `TM_FLOOR` | 6 | implemented | `—` | `fence header writer (single-sided)` |
| `TM_CEILING` | 7 | implemented | `—` | `fence header writer (single-sided)` |
| `TM_OBRACK` | 8 | declared-only | `—` | `—` |
| `TM_INTERVAL` | 9 | declared-only | `—` | `—` |
| `TM_ROOT` | 10 | implemented | `writeSqrtHeader/writeNthRootHeader` | `writeSqrtNode` |
| `TM_FRACT` | 11 | implemented | `writeFractionHeader` | `writeFractionNode` |
| `TM_UBAR` | 12 | implemented | `writeUnderlineHeader` | `writeCommandNode` |
| `TM_OBAR` | 13 | implemented | `writeOverlineHeader` | `writeCommandNode` |
| `TM_ARROW` | 14 | declared-only | `—` | `—` |
| `TM_INTEG` | 15 | implemented | `writeIntegralHeader` | `writeBigOpHeader (integrals)` |
| `TM_SUM` | 16 | implemented | `writeSumHeader` | `writeBigOpHeader (sum-like fallback)` |
| `TM_PROD` | 17 | implemented | `writeProductHeader` | `writeBigOpHeader (prod)` |
| `TM_COPROD` | 18 | declared-only | `—` | `—` |
| `TM_UNION` | 19 | declared-only | `—` | `—` |
| `TM_INTER` | 20 | declared-only | `—` | `—` |
| `TM_INTOP` | 21 | declared-only | `—` | `—` |
| `TM_SUMOP` | 22 | declared-only | `—` | `—` |
| `TM_LIM` | 23 | declared-only | `—` | `—` |
| `TM_HBRACE` | 24 | implemented | `writeHBraceHeader` | `writeHorizontalBrace (\overbrace / \underbrace)` |
| `TM_HBRACK` | 25 | implemented | `writeHBrackHeader` | `writeHorizontalBracket (\overbracket / \underbracket)` |
| `TM_LDIV` | 26 | implemented | `writeLongDivisionHeader` | `writeLongDivisionNode` |
| `TM_SUB` | 27 | implemented | `writeSubscriptHeader` | `writeSubscriptNode / writeSupSubAttachment` |
| `TM_SUP` | 28 | implemented | `writeSuperscriptHeader` | `writeSuperscriptNode / writeSupSubAttachment` |
| `TM_SUBSUP` | 29 | implemented | `writeSubSuperscriptHeader` | `writeSuperscriptNode / writeSupSubAttachment` |
| `TM_DIRAC` | 30 | declared-only | `—` | `—` |
| `TM_VEC` | 31 | implemented | `writeVecHeader` | `writeCommandNode` |
| `TM_TILDE` | 32 | implemented | `writeTildeHeader` | `writeCommandNode` |
| `TM_HAT` | 33 | implemented | `writeHatHeader` | `writeCommandNode` |
| `TM_ARC` | 34 | declared-only | `—` | `—` |
| `TM_JSTATUS` | 35 | declared-only | `—` | `—` |
| `TM_STRIKE` | 36 | implemented | `writeStrikeHeader` | `writeCommandNode` (`\\cancel` / `\\bcancel` / `\\xcancel`) |
| `TM_BOX` | 37 | implemented | `writeBoxHeader` | `writeCommandNode` (`\\boxed`) |
| `TM_LSCRIPT` | 44 | declared-only | `—` | `—` |

## Notes
- `implemented` means the repository contains both a template builder helper and an explicit writer path using that template family.
- `builder-only` means helper code exists but no explicit writer use was mapped in this first pass.
- `declared-only` means the official template constant exists in `MtefRecord`, but no builder/writer path was found yet.
- This matrix measures **code-path existence**, not semantic completeness. For example, integrals may be implemented while still missing full variation coverage (double/triple/loop).
