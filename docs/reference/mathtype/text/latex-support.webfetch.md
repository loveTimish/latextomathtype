# Web-fetch capture: MathType LaTeX support

Source: https://docs.wiris.com/en_US/mathtype-web-interface-features/latex-support
Captured via OpenClaw `web_fetch` because direct scripted TLS fetch was unstable from this Linux host.

## Key excerpts

- MathType integrations let authors edit formulas visually while storing/recovering LaTeX.
- During processing, MathType integrations convert LaTeX into **MathML**, and preserve the original LaTeX in `<semantics><annotation encoding="LaTeX">...`.
- The official support boundary is explicit:
  - **Not all LaTeX is supported**.
  - MathType supports the LaTeX instructions **which have a MathML equivalent**.
- Official references point to:
  - the example gallery
  - the `latex-coverage` command list page

## Architectural consequence

The repository should treat **MathML-equivalent LaTeX** as the primary supported subset and reserve direct-MTEF custom paths only for structures that MathType supports but that are awkward or non-standard in generic MathML.
