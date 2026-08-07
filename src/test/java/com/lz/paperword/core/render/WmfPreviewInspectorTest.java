package com.lz.paperword.core.render;

import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

class WmfPreviewInspectorTest {

    private final LaTeXImageRenderer renderer = new LaTeXImageRenderer();

    @Test
    void acceptsReadableUnicodeAndNestedStructurePreviews() {
        assertReadable("\\begin{array}{r}2.30€\\\\中文𝛼\\end{array}");
        assertReadable("\\frac{\\sqrt{a^{2}+b^{2}}}{\\sum_{i=1}^{n}x_i}");
        assertReadable("\\xcancel{0=1}");
    }

    @Test
    void rejectsCorruptedPlaceableChecksum() {
        byte[] wmf = renderer.renderForOlePreview("x+1").data().clone();
        wmf[20] ^= 0x01;

        WmfPreviewInspector.Inspection inspection = WmfPreviewInspector.inspect(wmf);

        assertFalse(inspection.valid());
        assertTrue(inspection.error().contains("checksum"));
    }

    private void assertReadable(String latex) {
        LaTeXImageRenderer.PreviewImage preview = renderer.renderForOlePreview(latex);
        WmfPreviewInspector.Inspection inspection = WmfPreviewInspector.inspect(preview.data());

        assertFalse(preview.placeholder(), latex);
        assertTrue(inspection.valid(), () -> latex + ": " + inspection.error());
        assertTrue(inspection.foregroundPixels() > 0, latex);
        assertFalse(inspection.inkTouchesEdge(), latex);
        assertTrue(inspection.foregroundDensity() <= 0.80d, latex);
    }
}
