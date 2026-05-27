# MathType / WIRIS local reference cache

This directory is the local mirror and working notebook for the official MathType/WIRIS references used by this repository.

## Goal

Keep the implementation anchored to official behavior instead of memory or folklore:

- MTEF v5 structure and template semantics
- OLE storage layout for `Equation Native`
- MathType's published LaTeX support boundary
- MathType's published MathML coverage boundary
- Example and coverage pages that can act as regression corpora

## Directory layout

- `raw/` — best-effort fetched HTML/XML snapshots
- `text/` — extracted text or manually curated web-fetch captures
- `manifest.txt` — fetch status for the scripted mirror pass
- `fetch_official_docs.py` — reproducible fetch script (best effort; some hosts are TLS-fragile)

## Current baseline sources

1. `mtef-v5` — official MTEF v5 reference
2. `mtef-storage` — how MTEF is embedded in OLE / clipboard / files
3. `latex-support` — official statement that supported LaTeX is the subset with MathML equivalents
4. `mathml-coverage` — official MathML coverage page
5. `latex-coverage` — WIRIS demo page enumerating supported commands/examples
6. `examples` — MathType SDK example gallery
7. `sitemap` — discovery index for future pulls

## Important working rule

For architecture and implementation decisions, prefer this precedence order:

1. Official WIRIS/MathType reference pages mirrored here
2. Verified behavior from current repository probes/tests
3. External community examples only as supplementary evidence

## Known limitation

The scripted fetch can fail on some WIRIS endpoints because of intermittent TLS issues from this Linux environment. When that happens, the fallback is to save the successful `web_fetch` extraction into `text/*.webfetch.*` files so the reference still exists locally.
