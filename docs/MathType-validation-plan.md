# MathType implementation validation plan

This file defines the acceptance gates for each implementation phase.

## Validation layers

### 1. Parse / normalization correctness

Artifacts:
- LaTeX input sample
- normalized IR dump (or current AST dump before IR lands)
- unsupported/fallback diagnostics

Pass criteria:
- no silent loss of structure
- unsupported constructs are explicit
- supported constructs normalize deterministically

### 2. MTEF structural correctness

Artifacts:
- binary assertions on template selector / variation / slot order
- record-level checks for LINE / TMPL / MATRIX / PILE / END balance

Pass criteria:
- expected template selector appears
- expected variation bits appear
- slot order matches MathType official docs / verified binary samples

### 3. Visual rendering correctness

Artifacts:
- preview PNG/SVG generated from the repository
- regression sample gallery

Pass criteria:
- formula is visually sane
- 2D layouts do not collapse into linear text
- fences, limits, scripts, and alignment remain readable

### 4. Editable-object correctness

Artifacts:
- generated `.docx`
- embedded OLE / `Equation Native` stream inspection
- fallback Linux checks when Word/MathType are unavailable

Pass criteria:
- OLE package is structurally valid
- `Equation Native` length/header is consistent
- proxy round-trip checks via extraction / conversion heuristics stay sane

## Linux fallback strategy

When local Word + MathType are unavailable, validation must still proceed by proxy:

1. Inspect generated OLE / POIFS structure directly
2. Assert MTEF template/variation bytes against official docs
3. Use preview rendering and sample comparison as visual proxy
4. Where useful, compare semantic structure against `docx2tex` / `docxtolatex` outputs instead of GUI editability
5. Keep a backlog of cases that still require a future Windows confirmation pass

## Sample buckets

### Core linear
- variables, numbers, operators, functions, greek letters

### Core scripted
- `x^2`, `x_i`, `x_i^2`, scripted fences

### Core structural
- fractions, roots, n-th roots, limits, big operators

### 2D standard
- matrix, pmatrix, bmatrix, cases, aligned, eq arrays

### K12 extended
- decimal vertical arithmetic, multiplication arrays, long division, cross multiplication

## Phase gate rule

Do not advance to the next phase until:

- code exists
- targeted tests/probes exist
- validation artifacts are saved or reproducible
- any remaining uncertainty is written down explicitly
