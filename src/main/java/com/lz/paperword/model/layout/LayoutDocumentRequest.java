package com.lz.paperword.model.layout;

import lombok.Data;

import java.util.ArrayList;
import java.util.List;

/**
 * PDF 版面导出接口使用的布局模型。
 */
@Data
public class LayoutDocumentRequest {

    private DocumentInfo document = new DocumentInfo();
    private List<Page> pages = new ArrayList<>();
    private FidelityReport fidelityReport;

    @Data
    public static class DocumentInfo {
        private String title;
        private String sourceFileName;
        private String sourceType = "pdf";
        private Integer pageWidth = 1000;
        private Integer pageHeight = 1414;
        private Integer marginTopTwips = 1440;
        private Integer marginBottomTwips = 1440;
        private Integer marginLeftTwips = 1440;
        private Integer marginRightTwips = 1440;
    }

    @Data
    public static class Page {
        private Integer pageNumber;
        private Integer width = 1000;
        private Integer height = 1414;
        private String backgroundImagePath;
        private List<QuestionMetadata> questions = new ArrayList<>();
        private List<Block> blocks = new ArrayList<>();
        private List<String> warnings = new ArrayList<>();
    }

    @Data
    public static class Block {
        private String id;
        private BlockType type = BlockType.PARAGRAPH;
        private BoundingBox bbox;
        private Style style = new Style();
        private String text;
        private String latex;
        private String imagePath;
        private String imageContentType;
        private Integer imageWidthPx;
        private Integer imageHeightPx;
        private List<TableRow> rows = new ArrayList<>();
        private List<String> warnings = new ArrayList<>();
    }

    public enum BlockType {
        TITLE,
        HEADING,
        PARAGRAPH,
        FORMULA,
        TABLE,
        IMAGE
    }

    @Data
    public static class BoundingBox {
        private Integer x;
        private Integer y;
        private Integer width;
        private Integer height;
    }

    @Data
    public static class Style {
        private Alignment alignment = Alignment.LEFT;
        private String fontFamily = "宋体";
        private Integer fontSizePt = 11;
        private boolean bold;
        private boolean italic;
        private Integer indentLeftTwips;
        private Integer spacingBeforeTwips;
        private Integer spacingAfterTwips;
    }

    public enum Alignment {
        LEFT,
        CENTER,
        RIGHT,
        BOTH
    }

    @Data
    public static class TableRow {
        private List<TableCell> cells = new ArrayList<>();
    }

    @Data
    public static class TableCell {
        private String text;
        private String backgroundColor;
        private Integer widthPct;
    }

    @Data
    public static class FidelityReport {
        private Integer score;
        private Integer totalPages;
        private Integer totalBlocks;
        private Integer totalQuestions;
        private Integer formulaBlocks;
        private Integer imageBlocks;
        private Integer tableBlocks;
        private Integer fallbackPages;
        private List<String> warnings = new ArrayList<>();
    }

    @Data
    public static class QuestionMetadata {
        private String questionId;
        private String serialNumber;
        private String questionType;
        private String stem;
        private List<OptionMetadata> options = new ArrayList<>();
        private String answer;
        private String analysis;
        private String analysisLabel;
        private String solution;
        private String difficulty;
        private String knowledgePoint;
        private List<String> latexList = new ArrayList<>();
        private List<String> tags = new ArrayList<>();
    }

    @Data
    public static class OptionMetadata {
        private String label;
        private String content;
    }
}
