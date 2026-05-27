# Web-fetch capture: How MTEF is stored in files and objects

Source: https://docs.wiris.com/en_US/mathtype-sdk/how-mtef-is-stored-in-files-and-objects
Captured via OpenClaw `web_fetch` because direct scripted TLS fetch was unstable from this Linux host.

## Key excerpts

- MathType stores MTEF in multiple wrappers so equations can be reopened later.
- OLE equation objects store MTEF as native data with a **28-byte header** followed by MTEF bytes.
- The documented OLE header fields are:
  - `cbHdr`
  - `version`
  - `cf` (`RegisterClipboardFormat("MathType EF")`)
  - `cbObject`
  - `reserved1..reserved4`
- Translator outputs (including LaTeX translators) can embed MTEF as plain text using the `MathType!MTEF!` convention.
- Text-mode MTEF uses a 64-character encoding alphabet and checksum markers.

## Why this matters here

This repository's `OlePackager` and `MathTypeEmbedder` should be validated against these documented storage rules, not only against ad-hoc sample files.
