# MathML-equivalent LaTeX subset roadmap

## Product goal

Support the LaTeX subset that MathType officially treats as convertible to/from MathML, while preserving repository-specific direct-MTEF paths for educational layouts that do not fit cleanly into generic MathML.

## Architecture stance

- **Primary semantic target:** MathML-equivalent math structures
- **Internal representation:** a MathML-aligned IR (not necessarily serialized XML at runtime)
- **Writer target:** MTEF/OLE for Word embedding
- **Fallback policy:** explicit diagnostics, never silent semantic loss

## Phases

### Phase 0 — Official reference localization
Goal:
- mirror critical MathType/WIRIS references locally
- write down source precedence and local discovery notes

Acceptance:
- local docs present in `docs/reference/mathtype/`
- fetch script exists
- failed online fetches are captured as curated local notes

Status:
- complete

### Phase 1 — Support matrix and validation regime
Goal:
- define coverage categories and current code-path status
- define objective validation gates

Acceptance:
- `docs/MathType-support-matrix.md` exists
- `docs/MathType-validation-plan.md` exists
- matrix generation is reproducible from source code

Status:
- complete

### Phase 2 — Validation backbone actually runs
Goal:
- make repository tests execute for real under Maven/JUnit 5
- remove environment-specific hard failures that block Linux CI

Acceptance:
- `mvn test` discovers and executes the suite
- no unconditional Windows-path failure remains

Status:
- complete (`93` tests executed; `0` failures/errors; `3` skips)

### Phase 3 — Introduce MathML-aligned IR
Goal:
- stop binding `MtefWriter` directly to ad-hoc LaTeX parsing shapes
- normalize supported LaTeX subset into canonical math structures

Initial IR categories:
- tokens: ident / number / operator / text
- one-dimensional: frac / sqrt / root / sub / sup / subsup / under / over / underover
- fencing: generalized fence / delimiter
- two-dimensional: table / row / cell
- enclosure: box / strike / overbrace / underbrace (as structure, not merely command text)
- extensions: long-division / vertical-arithmetic / cross-multiplication / unsupported

Acceptance:
- converter exists for currently supported core subset
- IR dumps can be generated for probes/tests
- unsupported nodes are explicit

Status:
- not started

### Phase 4 — MTEF template family completion
Goal:
- implement missing official template families and variations in writer

Priority family list:
1. limits and big-op distinctions (`TM_LIM`, integrals variations)
2. fence variants (`TM_DBAR`, `TM_FLOOR`, `TM_CEILING`, `TM_INTERVAL`, asymmetric fences)
3. 2D standard semantics (`cases`, `aligned`, matrix fence variants)
4. enclosures (`TM_BOX`, `TM_STRIKE`, `TM_HBRACE`, `TM_HBRACK`, arc)
5. advanced/rare templates (`TM_DIRAC`, `TM_LSCRIPT`, arrow families)

Acceptance:
- each family ships with binary structural assertions and representative examples

Status:
- in progress conceptually; not yet started as a systematic phase

### Phase 5 — Cross-validation without local Word/MathType
Goal:
- validate structure and output quality from Linux using proxies

Methods:
- inspect OLE/POIFS structure directly
- inspect `Equation Native` header + MTEF payload
- assert template selector / variation bytes
- render previews and compare visual sanity
- compare semantic structure with `docx2tex` / `docxtolatex` where useful

Acceptance:
- each new feature lands with a Linux-reproducible verification path

Status:
- partial; existing tests generate DOCX previews, but no unified probe/report script yet

## Current rule of engagement

Do not move a feature family forward unless:
1. supported shape is documented
2. unsupported behavior is explicit
3. validation artifact is reproducible on Linux
4. remaining uncertainty is written down in docs
