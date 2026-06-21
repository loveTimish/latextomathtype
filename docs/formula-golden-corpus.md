# Formula Golden Corpus

This corpus is the first stable acceptance layer for the self-written WMF
formula box model. It is not a MathType pixel clone target. It defines the
formula structures that must be shaped consistently before visual tuning.

## Structure Families

| Family | Purpose |
| --- | --- |
| `linear` | Single-baseline formulas such as `a+b=c`, `F=ma`. |
| `script` | Superscript/subscript chains, including short geometry labels. |
| `fraction` | Display-style stacked fractions with one bar. |
| `inline_fraction` | Compact fractions embedded in a text-like line. |
| `nested_fraction` | Fractions containing fractions above or below the bar. |
| `sqrt_simple` | One radical over linear content. |
| `sqrt_script` | Radical body contains scripts. |
| `sqrt_fraction` | Radical body contains a fraction. |
| `fraction_sqrt` | Fraction numerator or denominator contains a radical. |
| `sqrt_nested` | Radical body contains another radical. |
| `sqrt_nested_fraction` | Nested radical plus fraction content. |
| `large_operator` | Sum, integral, limit, and similar operator layouts. |
| `matrix_array` | Arrays, cases, matrices, or aligned systems. |
| `chemistry` | Chemical subscripts, ions, and reaction arrows. |
| `geometry_label` | Triangle/segment/area labels common in K12 geometry. |
| `text_mixed` | Formulas mixed with CJK or long text fragments. |

## Acceptance Rules

The machine-readable corpus lives at
`src/test/resources/formula-golden-corpus.tsv`.

Each `required` case must render through `VectorWmfFormulaRenderer` as
self-written vector WMF. The record-level gate checks:

- `ExtTextOut` records exist for glyph runs.
- `Polyline` records exist for structural lines when the family needs bars,
  radical signs, boxes, or accents.
- `StretchDIB` and bitmap records do not appear.
- The WMF window extent matches the requested physical box.
- Text and structure coordinates stay inside the physical box.

Each `optional` case is a known expansion candidate. If it renders, it must
meet the same vector WMF safety rules. If it does not render, the report records
it as a gap instead of failing the suite.

## Current Design Bias

Formula size starts from one standard glyph ladder:

- Main Latin math glyphs use Times New Roman italic for math variables.
- Script glyphs are smaller by a fixed visual ratio.
- Fractions, roots, and arrays compute boxes from child dimensions, then add
  structure-specific padding and line thickness.
- Spacing is structural. Do not insert visible or invisible spaces to fake
  width.
- Nested roots are their own family because each radical sign needs a readable
  checkmark, top bar, and body offset at its depth.
