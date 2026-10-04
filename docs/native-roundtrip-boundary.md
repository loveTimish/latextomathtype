# Native MathType round-trip boundary

The export pipeline can create editable MathType OLE data and a preview for a DOCX. That does not establish how another editor will regenerate the preview or change its frame when a user opens an equation and saves it.

## Separate three questions

1. **Native data**: Did the complete Equation Native stream and the other compound-file streams change?
2. **Display inputs**: Did the preview's meaningful drawing/state records, the Word shape dimensions, baseline/run position, crop, opacity, hidden state or style/settings context change?
3. **Observed rendering and editing**: Did a specified editor with a specified font environment actually show the same pixels and remain editable through the user's save operation?

The first two questions can be investigated read-only on Linux with the [round-trip inspector](mathtype-roundtrip-inspector.md). The third requires direct editor evidence; the inspector reports pixel comparison as not measured. It does not claim to simulate Word/WPS/MathType or perform a native save.

Equal native bytes do not imply equal previews or frames. A native editor may retain the equation data while rebuilding a previously generated preview, changing physical dimensions or baseline. Conversely, different WMF bytes can consist solely of specified unused font-array tail bytes or text-alignment padding. CFB container bytes may change only because storage timestamps or allocation layout changed.

Compare original → first save and first save → second save separately. A stable second save does not erase a first-save change, and one stable document does not establish behavior for all expressions or editor versions. Preserve the original and saved snapshots rather than rewriting either into a purported stable reference.

## Explicit non-goals

- No preview freezing, conversion to a static picture, macro injection or automatic corrective resave
- No Windows-only rendering prerequisite or replacement with OMML
- No semantic proof from normalized MTEF signatures, raw OLE hashes or image hashes
- No claim that a font/size/coordinate/pen/brush change is merely harmless metadata
- No guarantee of pixel equality when rasterization has not been performed

Unknown records, incomplete correspondence, malformed/truncated inputs and unsupported formats remain inconclusive. Reports must distinguish a verified byte/layout change from an unmeasured visual or mathematical conclusion. Exact native equality is useful evidence, but it is only one layer of the round-trip boundary.
