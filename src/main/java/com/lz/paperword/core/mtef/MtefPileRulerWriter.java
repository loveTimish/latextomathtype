package com.lz.paperword.core.mtef;

import com.lz.paperword.core.latex.LaTeXNode;
import com.lz.paperword.core.layout.VerticalLayoutSpec;
import com.lz.paperword.core.layout.VerticalLayoutSpec.TabStopKind;
import com.lz.paperword.core.layout.VerticalLayoutSpec.VerticalRow;
import com.lz.paperword.core.layout.VerticalLayoutSpec.VerticalSegment;
import com.lz.paperword.core.layout.VerticalLayoutSpec.VerticalTabStop;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.List;

/**
 * 竖式专用的 PILE/RULER MTEF writer。
 */
public class MtefPileRulerWriter {

    @FunctionalInterface
    public interface NodeSerializer {
        void write(ByteArrayOutputStream out, LaTeXNode node) throws IOException;
    }

    public void writeLayout(ByteArrayOutputStream out, VerticalLayoutSpec spec, NodeSerializer serializer) throws IOException {
        out.write(MtefRecord.PILE);
        int options = shouldWriteRuler(spec) ? MtefRecord.OPT_LP_RULER : 0x00;
        out.write(options);
        out.write(resolvePileHalign(spec));
        out.write(resolvePileValign(spec));
        if (shouldWriteRuler(spec)) {
            writeRuler(out, spec.tabStops());
        }
        if (spec.isLongDivision()) {
            writeLongDivisionHeaderLines(out, spec, serializer);
        }
        for (int rowIndex = 0; rowIndex < spec.rows().size(); rowIndex++) {
            VerticalRow row = spec.rows().get(rowIndex);
            if (spec.isLongDivision() && row.kind() == VerticalLayoutSpec.RowKind.DIVIDEND) {
                continue;
            }
            if (spec.kind() == VerticalLayoutSpec.Kind.DECIMAL) {
                writePlainDecimalLine(out, row, hasRuleBelowRow(spec, rowIndex), serializer);
            } else {
                writeSegmentsLine(out, resolveWriterSegments(spec, row, rowIndex), serializer);
            }
        }
        out.write(MtefRecord.END);
    }

    public void writePileLine(ByteArrayOutputStream out, VerticalRow row, NodeSerializer serializer) throws IOException {
        writeSegmentsLine(out, toWriterSegments(row.segments(), false), serializer);
    }

    private void writeLongDivisionHeaderLines(ByteArrayOutputStream out, VerticalLayoutSpec spec, NodeSerializer serializer) throws IOException {
        int anchorColumn = Math.max(spec.anchorColumn(), 2);
        if (!spec.longDivisionHeader().quotient().isBlank()) {
            writeSegmentsLine(out, List.of(new WriterSegment(spec.longDivisionHeader().quotient(), anchorColumn, SegmentStyle.PLAIN)), serializer);
        }
        writeSegmentsLine(out, List.of(
            new WriterSegment(spec.longDivisionHeader().divisor(), 0, SegmentStyle.PLAIN),
            new WriterSegment(")", 1, SegmentStyle.PLAIN),
            new WriterSegment(spec.longDivisionHeader().dividend(), anchorColumn, SegmentStyle.OVERLINE)
        ), serializer);
    }

    private void writeSegmentsLine(ByteArrayOutputStream out, List<WriterSegment> segments, NodeSerializer serializer) throws IOException {
        out.write(MtefRecord.LINE);
        out.write(0x00);

        int currentStop = -1;
        if (segments.isEmpty()) {
            out.write(MtefRecord.END);
            return;
        }
        for (WriterSegment segment : segments) {
            currentStop = writeTabAlignedSegment(out, segment, currentStop, serializer);
        }
        // 关闭最后一个 tab group，确保最后一段也落到对应制表位。
        writeTabChar(out);
        out.write(MtefRecord.END);
    }

    private void writePlainDecimalLine(ByteArrayOutputStream out, VerticalRow row, boolean underline,
                                       NodeSerializer serializer) throws IOException {
        out.write(MtefRecord.LINE);
        out.write(0x00);
        String text = buildDecimalText(row);
        if (!text.isBlank()) {
            LaTeXNode lineNode = underline
                ? buildSegmentNode(new WriterSegment(text, 0, SegmentStyle.UNDERLINE))
                : buildGroup(text);
            serializer.write(out, lineNode);
        }
        out.write(MtefRecord.END);
    }

    public int writeTabAlignedSegment(ByteArrayOutputStream out, WriterSegment segment, int currentStop,
                                      NodeSerializer serializer) throws IOException {
        int targetStop = Math.max(segment.endColumn(), currentStop);
        while (currentStop < targetStop) {
            writeTabChar(out);
            currentStop++;
        }
        serializer.write(out, buildSegmentNode(segment));
        return currentStop;
    }

    public void writeRuler(ByteArrayOutputStream out, List<VerticalTabStop> tabStops) throws IOException {
        out.write(MtefRecord.RULER);
        out.write(tabStops.size() & 0xFF);
        for (VerticalTabStop tabStop : tabStops) {
            out.write(encodeTabStopType(tabStop.kind()) & 0xFF);
            writeInt16(out, tabStop.offset());
        }
    }

    private LaTeXNode buildSegmentNode(WriterSegment segment) {
        if (segment.style() == SegmentStyle.UNDERLINE) {
            LaTeXNode command = new LaTeXNode(LaTeXNode.Type.COMMAND, "\\underline");
            command.addChild(buildGroup(segment.text()));
            return command;
        }
        if (segment.style() == SegmentStyle.OVERLINE) {
            LaTeXNode command = new LaTeXNode(LaTeXNode.Type.COMMAND, "\\overline");
            command.addChild(buildGroup(segment.text()));
            return command;
        }
        return buildGroup(segment.text());
    }

    private LaTeXNode buildGroup(String text) {
        LaTeXNode group = new LaTeXNode(LaTeXNode.Type.GROUP);
        if (text == null) {
            return group;
        }
        for (char ch : text.toCharArray()) {
            group.addChild(new LaTeXNode(LaTeXNode.Type.CHAR, String.valueOf(ch)));
        }
        return group;
    }

    private void writeTabChar(ByteArrayOutputStream out) throws IOException {
        writeCharRecord(out, MtefRecord.FN_TEXT, '\t');
    }

    private void writeCharRecord(ByteArrayOutputStream out, int typeface, int mtcode) throws IOException {
        out.write(MtefRecord.CHAR);
        out.write(0x00);
        out.write(encodeTypeface(typeface));
        out.write(mtcode & 0xFF);
        out.write((mtcode >> 8) & 0xFF);
    }

    private void writeInt16(ByteArrayOutputStream out, int value) throws IOException {
        out.write(value & 0xFF);
        out.write((value >> 8) & 0xFF);
    }

    private int encodeTypeface(int typeface) {
        return (typeface & 0x7F) | 0x80;
    }

    private int encodeTabStopType(TabStopKind kind) {
        return switch (kind) {
            case LEFT -> 0;
            case CENTER -> 1;
            case RIGHT -> 2;
            case RELATION -> 3;
            case DECIMAL -> 4;
        };
    }

    private int resolvePileValign(VerticalLayoutSpec spec) {
        if (spec.kind() == VerticalLayoutSpec.Kind.DECIMAL) {
            // 参考小数点对齐对象使用的是 decimal halign + center-line valign 组合。
            return 1;
        }
        return spec.isLongDivision() ? 2 : 2;
    }

    private int resolvePileHalign(VerticalLayoutSpec spec) {
        // 小数加减法按用户要求改成整行右对齐，不再使用 decimal 对齐语义。
        return spec.kind() == VerticalLayoutSpec.Kind.DECIMAL ? 3 : 1;
    }

    private List<WriterSegment> resolveWriterSegments(VerticalLayoutSpec spec, VerticalRow row, int rowIndex) {
        boolean underline = hasRuleBelowRow(spec, rowIndex);
        if (spec.kind() == VerticalLayoutSpec.Kind.DECIMAL) {
            return resolveDecimalSegments(spec, row, underline);
        }
        return toWriterSegments(row.segments(), underline);
    }

    private List<WriterSegment> resolveDecimalSegments(VerticalLayoutSpec spec, VerticalRow row, boolean underline) {
        String text = buildDecimalText(row);
        if (text.isBlank()) {
            return List.of();
        }
        SegmentStyle style = underline ? SegmentStyle.UNDERLINE : SegmentStyle.PLAIN;
        return List.of(new WriterSegment(text, Math.max(spec.decimalColumn(), 0), style));
    }

    private boolean shouldWriteRuler(VerticalLayoutSpec spec) {
        // 参考小数点对齐对象只使用 decimal pile，不带 ruler/tab。
        return spec.kind() != VerticalLayoutSpec.Kind.DECIMAL && !spec.tabStops().isEmpty();
    }

    private String buildDecimalText(VerticalRow row) {
        StringBuilder builder = new StringBuilder();
        for (String cell : row.cells()) {
            if (cell != null && !cell.isBlank()) {
                builder.append(cell);
            }
        }
        return builder.toString();
    }

    private boolean hasRuleBelowRow(VerticalLayoutSpec spec, int rowIndex) {
        int boundaryIndex = rowIndex + 1;
        return spec.ruleSpans().stream().anyMatch(span -> span.boundaryIndex() == boundaryIndex);
    }

    private List<WriterSegment> toWriterSegments(List<VerticalSegment> segments, boolean underline) {
        return segments.stream()
            .map(segment -> new WriterSegment(
                segment.text(),
                segment.endColumn(),
                (underline || segment.underlined()) ? SegmentStyle.UNDERLINE : SegmentStyle.PLAIN))
            .toList();
    }

    private enum SegmentStyle {
        PLAIN,
        UNDERLINE,
        OVERLINE
    }

    private record WriterSegment(String text, int endColumn, SegmentStyle style) {
    }
}
