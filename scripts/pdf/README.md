# Explicit Linux PDF baseline compatibility export

This helper creates **two genuine LibreOffice PDF exports from the same imported DOCX**:

- `*-native.pdf`: normal LibreOffice layout before any compatibility adjustment
- `*-baseline-compatible.pdf`: the same model after restoring the inline OLE baseline that LibreOffice's DOCX importer discards

It never overwrites/saves the source DOCX, changes equation content, replaces an OLE with a picture, or edits PDF pixels. All MathType objects are traversed in actual body/table order, matched to DOCX object count and dimensions, and reported in a TSV audit. Unsupported mappings fail explicitly; objects are never selectively skipped.

## Why a separate path exists

WordprocessingML `w:position` is the normal run baseline offset. LibreOffice's `appendOLE()` inserts OLEs without those run properties; its VML replacement-image path forces inline `VertOrient=TOP`, aligning the object's bottom to the text baseline.

The DOCX exporter now derives its baseline from the actual MathJax SVG through the exact Batik/WMF transform. This helper restores that DOCX metric in the LO working model using supported UNO APIs:

```
descent_hmm = -w_position_half_points * 2540 / 144
VertOrient = NONE
VertOrientPosition = -Height - TopMargin + descent_hmm
AnchorType remains AS_CHARACTER
```

Units are 1/100 mm. Top/bottom margins, object dimensions, native OLE content and source DOCX remain unchanged. `BottomMargin` stays as line whitespace, not a baseline shift. Writer internally rounds to twips; the readback guard permits at most 2 hundredths of a millimetre.

This is **not a claim that the DOCX alone fixes every converter**. A plain `soffice --convert-to pdf` can still show the original importer limitation. Word/MathType native GUI editing remains a separate acceptance step.

## Requirements and use

- Java 21 with the compiler module
- Python 3.8 or later
- LibreOffice with the socket/URP UNO bridge
- Official public UNO classes `org.libreoffice:libreoffice:24.8.0`

Obtain the Java UNO jar through your normal Maven setup:

```
mvn dependency:copy -Dartifact=org.libreoffice:libreoffice:24.8.0 -DoutputDirectory=target/pdf-tools
```

The tested official Maven artifact is 2,240,422 bytes, SHA-1 `246252d31b3f03c721278783b445ad9ee2952347`.
Source: https://central.sonatype.com/artifact/org.libreoffice/libreoffice/24.8.0

```
python scripts/pdf/export_with_baseline.py input.docx output/new-run \
  --uno-jar target/pdf-tools/libreoffice-24.8.0.jar
```

Each run starts its own loopback-only LibreOffice process and temporary profile. No macro security settings are changed. The document is loaded hidden/read-only with macros disabled and link updates disabled. XML external entities, linked OLEs and external package relationships are rejected. The helper refuses to overwrite existing outputs.

Supported input: this project's embedded MathType, horizontal nonrotated inline OLEs in body paragraphs and simple tables. Header/footer/frame OLEs, rotated/flipped shapes, unsupported cell naming, missing baselines, foreign OLE types, count/order/dimension mismatches fail instead of being approximated.

The source SHA-256, both PDFs, process log and per-object baseline audit are retained. Verification should include page PNGs, `pdfinfo`, `pdfimages -list`, source/MTEF/WMF hashes, and visual inspection. In the tested LibreOffice 26.8 build, **both PDF paths rasterize OLE previews internally at about 300 dpi**; the original DOCX still contains vector WMF and editable OLE. Do not advertise these PDFs as all-vector.

## Primary implementation references

- Word `w:position`: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.position?view=openxml-3.0.1
- LO OLE import: https://github.com/LibreOffice/core/blob/master/sw/source/writerfilter/dmapper/DomainMapper_Impl.cxx
- LO VML import: https://github.com/LibreOffice/core/blob/master/oox/source/vml/vmlshape.cxx
- Writer inline positioning: https://github.com/LibreOffice/core/blob/master/sw/source/core/objectpositioning/ascharanchoredobjectposition.cxx
- Writer fly frame placement: https://github.com/LibreOffice/core/blob/master/sw/source/core/layout/flyincnt.cxx
- UNO frame properties: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1text_1_1BaseFrameProperties.html

## HTTP integration and automated lifecycle tests

See [Linux PDF service](../../docs/linux-pdf-service.md) for the opt-in `/api/export/pdf` and `/api/export/layout-pdf` routes, bounded concurrency, dependency errors, and controlled cleanup. Run `python scripts/pdf/test_export_with_baseline.py -v` for synthetic process-lifecycle checks. CLI output is staged until all assertions and the source hash check pass; failure leaves only a bounded diagnostic log, never a published partial PDF.

## Missing CJK fonts

The CLI defaults to `--cjk-font "Noto Serif CJK SC"`. After exporting the untouched native comparison, compatible output replaces only unavailable Asian font families on CJK text portions with this explicitly installed family. Available requested fonts are retained; source text, font sizes, OLEs and source DOCX do not change. UNO checks that the replacement exists and logs `CJK_FONT_FALLBACK` mappings/counts. Missing required replacement is a dependency error. Pass an empty value to explicitly disable this policy; that re-exposes native missing-font risks. Genuine replacement metrics may reflow pages. The six-question repository template needs this policy in the tested Linux environment to prevent two headings losing Chinese characters.
