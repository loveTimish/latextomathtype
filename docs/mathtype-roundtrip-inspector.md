# Read-only MathType DOCX round-trip inspector

`com.lz.paperword.core.mtef.MathTypeRoundTripInspectCli` compares two existing DOCX snapshots. It does not launch Word, WPS or MathType, execute macros, modify OLE objects, regenerate previews, or repair documents. No Windows or OMML conversion is involved.

## Run

After building the Java 21 project and obtaining its runtime dependency classpath:

```sh
java -cp "target/classes:$DEPENDENCY_CLASSPATH" \
  com.lz.paperword.core.mtef.MathTypeRoundTripInspectCli \
  before.docx after.docx new-report.json
```

The report path must not already exist. Output uses atomic create-new semantics and rejects input aliases, hardlinks and symlinks. Both DOCX inputs remain read-only. The CLI prints the result and counts. Exit codes are `0` for `OBSERVED_PRESERVED`, `1` for `OBSERVED_CHANGED`, and `2` for `INCONCLUSIVE` or an invocation/report-writing error. A preservation result remains a finite observation, not a fidelity guarantee. Consumers must inspect `status` and the individual layers.

Java API: `compare(Path, Path)`, `inspect(Path)` and `writeReport(report, output, inputs...)`. Reports contain hashes, record bytes and part identifiers. Paragraph text is used locally for correspondence but only its hash is reported. MTEF record hex can contain source text; treat a report with the same confidentiality as its input documents.

## Correspondence

The inspector discovers transitional Word `o:OLEObject` elements in Word XML parts. It follows the owning part's actual `.rels` entries to the OLE part, and follows `ShapeID` to that object's VML shape and image relationship. It never pairs formulas by embedding filename, ZIP order or numeric suffix.

Pairs use a unique stable shape ID, then a unique paragraph-text context, then the single-object case if each document contains exactly one object. Missing, duplicate or ambiguous correspondence remains inconclusive. Ordinals are for reporting and position checks, not correspondence. Unreferenced embeddings and unsupported/external relationships are reported.

## Independent observations

The JSON separates:

- Complete `Equation Native` stream SHA-256, including its OLE header and preferences
- Ordered, lossless MTEF header/record bytes, including font, size, color, coordinates, definitions, templates and END records
- Exact `EQN_PREFS` record bytes in their original order, separately from the other sequence bytes
- Every other CFB stream's name, size and hash, directory class IDs and root class ID
- CFB whole-container hash and root modified FILETIME, separately from stream content
- Complete shape, run, paragraph and ancestor-property context, including same-kind sibling positions, width, height, crop, baseline/run position, hidden state, opacity and style references
- Conservative Word document, styles, settings, numbering, theme and other Word XML context hashes
- Raw preview bytes and bounded WMF effective-record comparison

The inspector does **not** use `MtefRecordNormalizer` or its lossy canonical signatures as evidence of mathematical equivalence. Unknown future records are retained losslessly but keep interpretation inconclusive. Malformed native data is unknown, not evidence that the mathematics changed.

The narrow Word metadata exclusions are namespace declarations and equivalent prefix/attribute ordering, Word `rsid*`, `rsids`, proofing markers, `_GoBack` bookmarks, and gallery-only `qFormat` elements/attributes. Relationship IDs and VML/OLE instance IDs are resolved or replaced with their role; actual shape/run position is retained separately. Semantic text, including space-only `w:t` and instruction text, is preserved. Other XML differences are conservative changes; this can include changes that ultimately have no visible effect.

## WMF policy

WMF headers, declared sizes, record lengths, point counts, string lengths, optional `Dx` arrays, termination and bounds are checked. Known drawing/state records retain all their bytes. LOGFONT font height and other fields, font names up to the NUL, text bytes, every coordinate, `Dx`, pen, brush and color remain significant.

Exactly two byte ranges can be zeroed in the comparison copy:

1. `META_CREATEFONTINDIRECT`: bytes after the first NUL within the fixed 32-byte `FaceName` field
2. `META_EXTTEXTOUT`: its single alignment byte when `StringLength` is odd

All other bytes are retained, including reserved fields, comments and unknown records. SETBKMODE and SETTEXTALIGN accept their specified optional two-byte reserved field without deleting it. Historical MFCOMMENT/escape 15 framing is recognized and its complete opaque payload remains significant; embedded formats are not rendered or interpreted. Other unknown functions/options make the effective comparison inconclusive. Malformed/truncated records never pass.

“Effective records changed” means that retained drawing/state input bytes changed. It is **not a claim that pixels changed**: there is no complete WMF or Word renderer in this tool, and every report states `pixelComparison: NOT_MEASURED`. Equal bytes are likewise no guarantee of equal display across font installations, editor versions or hosts.

Specification references: [MS-WMF](https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-wmf/), [EXTTEXTOUT](https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-wmf/7d07c44a-a828-4b82-9af0-e0a81cced5a8), [SETTEXTALIGN](https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-wmf/562b5f06-dc3e-4446-bb2f-9ada932a8f9d), [MTEF v5](https://docs.wiris.com/en_US/mathtype-mtef-v5-mathtype-40-and-later).

## Conservative status meanings

- `OBSERVED_PRESERVED`: supported compared layers did not change within this finite observation policy; not a mathematical, pixel, editing or round-trip certification
- `OBSERVED_CHANGED`: at least one reliable compared layer changed; read the layer to determine whether this is native data, state/records, layout or context
- `INCONCLUSIVE`: no known significant change establishes the result, and correspondence, support or input integrity is incomplete

A known frame change remains changed even if MTEF interpretation is unknown. Conversely, equal native/frame/effective-WMF hashes can coexist with an overall inconclusive result because an opaque future MTEF record is unsupported. Raw WMF padding differences or CFB timestamp-only differences do not independently force an overall changed result. An unreadable package cannot establish object deletion and always produces an inconclusive result.

## Input limits and unsupported cases

ZIP limits are 4,096 entries, 16 MiB expanded per entry and 128 MiB total (plus a 128 MiB compressed file cap). The inspector rejects duplicate and unsafe ZIP paths, escaping internal relationship targets and unavailable targets. It never extracts ZIP entries to disk or opens external targets.

XML DTDs, external entities, XInclude and external schema/DTD access are disabled. XML nesting is limited to 64 and nodes to 500,000 per part. CFB sectors, FAT/DIFAT chains, property-tree cycles/depth, stream sizes, directory counts and cumulative stream bytes are bounded before/while POI reads streams. CFB limits are 4,096 directory entries, depth 64 and 16 MiB cumulative streams. MTEF/WMF record counts are bounded to 100,000; MTEF nesting is bounded to 64.

This is a finite transitional DOCX/VML MathType inspector, not a universal Office validator. Strict OOXML, alternative object markup, arbitrary preview formats, unknown WMF operations, malformed storage, and unsupported MTEF constructs may be inconclusive. Empty documents are not preservation evidence. Unit fixtures are synthesized in memory and contain no customer files.
