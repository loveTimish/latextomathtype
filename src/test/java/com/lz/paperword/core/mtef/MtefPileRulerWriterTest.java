package com.lz.paperword.core.mtef;

import com.lz.paperword.core.layout.VerticalLayoutSpec;
import com.lz.paperword.core.layout.VerticalLayoutSpec.RowKind;
import com.lz.paperword.core.layout.VerticalLayoutSpec.RuleSpan;
import com.lz.paperword.core.layout.VerticalLayoutSpec.TabStopKind;
import com.lz.paperword.core.layout.VerticalLayoutSpec.VerticalRow;
import com.lz.paperword.core.layout.VerticalLayoutSpec.VerticalSegment;
import com.lz.paperword.core.layout.VerticalLayoutSpec.VerticalTabStop;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayOutputStream;
import java.util.List;

import static org.junit.jupiter.api.Assertions.assertTrue;

class MtefPileRulerWriterTest {

    private final MtefPileRulerWriter writer = new MtefPileRulerWriter();

    @Test
    void shouldWriteRulerRecordWithTabStops() throws Exception {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        writer.writeRuler(out, List.of(
            new VerticalTabStop(0, TabStopKind.RIGHT, 240),
            new VerticalTabStop(1, TabStopKind.DECIMAL, 480)
        ));

        byte[] bytes = out.toByteArray();
        assertTrue(containsBytes(bytes, new byte[]{(byte) MtefRecord.RULER, 0x02}),
            "ruler header should contain record type and stop count");
        assertTrue(containsBytes(bytes, new byte[]{0x02, (byte) 0xF0, 0x00}),
            "right tab stop should encode type=2 and little-endian offset");
        assertTrue(containsBytes(bytes, new byte[]{0x04, (byte) 0xE0, 0x01}),
            "decimal tab stop should encode type=4 and little-endian offset");
    }

    @Test
    void shouldWritePileWithRulerAndTabCharacters() throws Exception {
        VerticalLayoutSpec spec = new VerticalLayoutSpec(
            VerticalLayoutSpec.Kind.ARITHMETIC,
            3,
            2,
            -1,
            List.of(
                new VerticalTabStop(0, TabStopKind.RIGHT, 240),
                new VerticalTabStop(1, TabStopKind.RIGHT, 480),
                new VerticalTabStop(2, TabStopKind.RIGHT, 720)
            ),
            List.of(
                new VerticalRow(RowKind.OPERAND, List.of("", "1", "2"),
                    List.of(new VerticalSegment("12", 1, 2, true))),
                new VerticalRow(RowKind.RESULT, List.of("", "", "3"),
                    List.of(new VerticalSegment("3", 2, 2, false)))
            ),
            List.of(new RuleSpan(1, 1, 2)),
            null
        );

        ByteArrayOutputStream out = new ByteArrayOutputStream();
        writer.writeLayout(out, spec, (buffer, node) -> {
            for (var child : node.getChildren()) {
                if (child.getValue() != null) {
                    buffer.write(MtefRecord.CHAR);
                    buffer.write(0x00);
                    buffer.write((MtefRecord.FN_TEXT & 0x7F) | 0x80);
                    buffer.write(child.getValue().charAt(0));
                    buffer.write(0x00);
                }
            }
        });

        byte[] bytes = out.toByteArray();
        assertTrue(containsBytes(bytes, new byte[]{(byte) MtefRecord.PILE, MtefRecord.OPT_LP_RULER, 0x01, 0x02}),
            "pile should include ruler option and halign/valign bytes");
        assertTrue(containsBytes(bytes, new byte[]{(byte) MtefRecord.RULER, 0x03}),
            "pile should embed a ruler record");
        assertTrue(containsBytes(bytes, new byte[]{(byte) MtefRecord.CHAR, 0x00, (byte) ((MtefRecord.FN_TEXT & 0x7F) | 0x80), 0x09, 0x00}),
            "tab characters should be serialized inside pile lines");
    }

    private boolean containsBytes(byte[] bytes, byte[] needle) {
        outer:
        for (int i = 0; i <= bytes.length - needle.length; i++) {
            for (int j = 0; j < needle.length; j++) {
                if (bytes[i + j] != needle[j]) {
                    continue outer;
                }
            }
            return true;
        }
        return false;
    }
}
