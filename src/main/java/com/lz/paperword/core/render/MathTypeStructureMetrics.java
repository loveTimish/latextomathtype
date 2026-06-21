package com.lz.paperword.core.render;

/**
 * Seed metrics derived from source MathType WMF structure reports.
 *
 * <p>The values mirror analysis/wmf-structure-metrics/
 * combined-xsc-tex-toggle-summary-parameter-candidates.txt. They are renderer
 * starting constants, not final visual acceptance gates.</p>
 */
final class MathTypeStructureMetrics {

    enum Family {
        LINEAR("linear"),
        SCRIPT("script"),
        SCRIPT_FRACTION("script_fraction"),
        SCRIPT_FRACTION_MIXED("script_fraction_mixed"),
        ORDINARY_FRACTION("fraction"),
        NESTED_FRACTION("nested_fraction"),
        TEXT_FRACTION("text_fraction"),
        SQRT("sqrt"),
        SQRT_NESTED("sqrt_nested"),
        SQRT_FRACTION("sqrt_fraction"),
        ARRAY("array"),
        ACCENT("accent");

        private final String previewClass;

        Family(String previewClass) {
            this.previewClass = previewClass;
        }

        String previewClass() {
            return previewClass;
        }
    }

    record FamilyMetrics(Family family, double heightPt, boolean sourceSeededHeight) {
        String previewClass() {
            return family.previewClass();
        }
    }

    record InkMetrics(double widthRatio, double heightRatio, double centerYRatio, int sampleCount) {
        boolean hasInkSamples() {
            return sampleCount > 0;
        }
    }

    record SourceSampleMetrics(
        Family family,
        double candidateHeightPt,
        double mainFontPt,
        double sourceScriptRatio,
        InkMetrics ink
    ) {
        boolean hasInkSamples() {
            return ink.hasInkSamples();
        }
    }

    static final double STANDARD_MAIN_FONT_PT = 10.495d;
    static final double SOURCE_SCRIPT_FONT_RATIO = 0.577d;
    static final double VISUAL_SCRIPT_FONT_RATIO = 0.78d;
    static final double STANDARD_SCRIPT_WIDTH_SCALE = 0.98d;
    static final double LINEAR_PREVIEW_WIDTH_SCALE = 1.12d;
    static final double SQRT_PREVIEW_WIDTH_SCALE = 1.18d;
    static final double SQRT_SCRIPT_PREVIEW_WIDTH_SCALE = 0.96d;
    static final double SQRT_FRACTION_PREVIEW_WIDTH_SCALE = 0.83d;
    static final double SQRT_FRACTION_MIXED_PREVIEW_WIDTH_SCALE = 1.08d;
    static final double SQRT_BODY_FRACTION_SCALE = 0.78d;
    static final double SQRT_BODY_FRACTION_LEFT_ADJUST_PT = -2.4d;
    static final double SQRT_BODY_FRACTION_TOP_PAD_PT = 2.4d;

    static final double SCRIPT_CANDIDATE_HEIGHT_PT = 18.75d;
    static final double ORDINARY_FRACTION_CANDIDATE_HEIGHT_PT = 30.75d;
    static final double NESTED_FRACTION_CANDIDATE_HEIGHT_PT = 60.0d;
    static final double LINEAR_HEIGHT_PT = 13.0d;
    static final double SYMBOL_LINEAR_HEIGHT_PT = 12.75d;
    static final double SCRIPT_HEIGHT_PT = 16.0d;
    static final double SCRIPT_FRACTION_HEIGHT_PT = 24.0d;
    static final double SCRIPT_FRACTION_MIXED_HEIGHT_PT = 33.75d;
    static final double ORDINARY_FRACTION_HEIGHT_PT = 28.0d;
    static final double COMPACT_INLINE_FRACTION_HEIGHT_PT = 25.5d;
    static final double SQRT_HEIGHT_PT = 18.0d;
    static final double SQRT_NESTED_HEIGHT_PT = 64.0d;
    static final double SQRT_FRACTION_HEIGHT_PT = 35.25d;
    static final double ORDINARY_FRACTION_ABOVE_PT = 16.0d;
    static final double ORDINARY_FRACTION_BELOW_PT = ORDINARY_FRACTION_HEIGHT_PT - ORDINARY_FRACTION_ABOVE_PT;
    static final double SQRT_FRACTION_ABOVE_PT = 20.0d;
    static final double SQRT_FRACTION_BELOW_PT = SQRT_FRACTION_HEIGHT_PT - SQRT_FRACTION_ABOVE_PT;
    static final double COMPACT_INLINE_FRACTION_ABOVE_PT = 13.0d;
    static final double COMPACT_INLINE_FRACTION_BELOW_PT =
        COMPACT_INLINE_FRACTION_HEIGHT_PT - COMPACT_INLINE_FRACTION_ABOVE_PT;
    static final double NESTED_FRACTION_HEIGHT_PT = 53.25d;
    static final double NESTED_FRACTION_ABOVE_PT = 30.0d;
    static final double NESTED_FRACTION_BELOW_PT = NESTED_FRACTION_HEIGHT_PT - NESTED_FRACTION_ABOVE_PT;
    static final double TEXT_FRACTION_HEIGHT_PT = 33.0d;
    static final double SQRT_BODY_LEFT_PAD_PT = 6.0d;
    static final double SQRT_BODY_FRACTION_LEFT_PAD_PT = 5.4d;
    static final double SQRT_BODY_Y_OFFSET_PT = 1.2d;
    static final double SQRT_NESTED_BODY_Y_EXTRA_PT = 3.0d;
    static final double SQRT_WIDTH_PAD_PT = 7.0d;
    static final double SQRT_CHECK_MID_X_PT = 2.0d;
    static final double SQRT_TALL_CHECK_LOW_X_PT = 1.1d;
    static final double SQRT_TALL_CHECK_SHOULDER_X_PT = 1.9d;
    static final double SQRT_TALL_CHECK_MID_X_PT = 1.4d;
    static final double SQRT_TALL_CHECK_LOW_Y_RATIO = 0.80d;
    static final double SQRT_TALL_CHECK_SHOULDER_Y_RATIO = 0.55d;
    static final double SQRT_CHECK_TOP_X_PT = 5.0d;
    static final double SQRT_TALL_CHECK_TOP_X_PT = 4.0d;
    static final double SQRT_TOP_Y_PT = 2.0d;
    static final double SQRT_BOTTOM_PAD_PT = 4.8d;
    static final double SQRT_TALL_BOTTOM_PAD_PT = 10.5d;
    static final double SQRT_TALL_EXTRA_HEIGHT_PT = 4.0d;
    static final double SQRT_LEFT_DESCENT_RATIO = 0.52d;
    static final double SQRT_BOX_TOP_OFFSET_PT = 1.3d;
    static final double SQRT_BOX_BODY_BASELINE_OFFSET_PT = 0.2d;
    static final double SQRT_BOX_LEFT_DESCENT_RATIO = 0.52d;
    static final double STRUCTURE_LINE_WIDTH_PT = 0.32d;
    static final double ARRAY_SOURCE_HEIGHT_PT = 33.0d;
    static final double ARRAY_ROW_HEIGHT_PT = 16.5d;
    static final double ARRAY_MIN_HEIGHT_PT = ARRAY_SOURCE_HEIGHT_PT;
    static final double ACCENT_HEIGHT_PT = 15.75d;
    static final double OVERLINE_Y_PT = 1.8d;
    static final double UNDERLINE_BOTTOM_PAD_PT = 1.2d;

    static boolean isTallSqrt(double rootHeightPt) {
        return rootHeightPt > SQRT_HEIGHT_PT + SQRT_TALL_EXTRA_HEIGHT_PT;
    }

    static double sqrtBottomPadPt(double rootHeightPt) {
        return isTallSqrt(rootHeightPt) ? SQRT_TALL_BOTTOM_PAD_PT : SQRT_BOTTOM_PAD_PT;
    }

    static SourceSampleMetrics sourceSampleMetrics(Family family) {
        double height = metrics(family).heightPt();
        return switch (family) {
            case LINEAR -> new SourceSampleMetrics(
                family,
                height,
                STANDARD_MAIN_FONT_PT,
                SOURCE_SCRIPT_FONT_RATIO,
                new InkMetrics(Double.NaN, Double.NaN, Double.NaN, 0)
            );
            case SCRIPT -> new SourceSampleMetrics(
                family,
                SCRIPT_CANDIDATE_HEIGHT_PT,
                12.0d,
                0.583d,
                new InkMetrics(0.922d, 0.711d, 0.487d, 3)
            );
            case SCRIPT_FRACTION -> new SourceSampleMetrics(
                family,
                SCRIPT_FRACTION_HEIGHT_PT,
                11.995d,
                0.583d,
                new InkMetrics(0.895d, 0.771d, 0.490d, 1)
            );
            case SCRIPT_FRACTION_MIXED -> new SourceSampleMetrics(
                family,
                SCRIPT_FRACTION_MIXED_HEIGHT_PT,
                12.0d,
                0.583d,
                new InkMetrics(0.944d, 0.838d, 0.522d, 1)
            );
            case ORDINARY_FRACTION -> new SourceSampleMetrics(
                family,
                ORDINARY_FRACTION_CANDIDATE_HEIGHT_PT,
                11.995d,
                Double.NaN,
                new InkMetrics(0.889d, 0.806d, 0.516d, 4)
            );
            case NESTED_FRACTION -> new SourceSampleMetrics(
                family,
                NESTED_FRACTION_CANDIDATE_HEIGHT_PT,
                11.997d,
                Double.NaN,
                new InkMetrics(0.911d, 0.871d, 0.522d, 2)
            );
            case TEXT_FRACTION -> new SourceSampleMetrics(
                family,
                height,
                STANDARD_MAIN_FONT_PT,
                Double.NaN,
                new InkMetrics(Double.NaN, Double.NaN, Double.NaN, 0)
            );
            case SQRT -> new SourceSampleMetrics(
                family,
                SQRT_HEIGHT_PT,
                12.0d,
                0.583d,
                new InkMetrics(0.881d, 0.750d, 0.486d, 5)
            );
            case SQRT_NESTED -> new SourceSampleMetrics(
                family,
                SQRT_NESTED_HEIGHT_PT,
                12.0d,
                0.583d,
                new InkMetrics(Double.NaN, Double.NaN, Double.NaN, 0)
            );
            case SQRT_FRACTION -> new SourceSampleMetrics(
                family,
                SQRT_FRACTION_HEIGHT_PT,
                11.995d,
                0.583d,
                new InkMetrics(0.885d, 0.871d, 0.507d, 3)
            );
            case ARRAY -> new SourceSampleMetrics(
                family,
                height,
                STANDARD_MAIN_FONT_PT,
                Double.NaN,
                new InkMetrics(Double.NaN, Double.NaN, Double.NaN, 0)
            );
            case ACCENT -> new SourceSampleMetrics(
                family,
                height,
                STANDARD_MAIN_FONT_PT,
                Double.NaN,
                new InkMetrics(Double.NaN, Double.NaN, Double.NaN, 0)
            );
        };
    }

    static FamilyMetrics metrics(Family family) {
        return metrics(family, -1L);
    }

    static FamilyMetrics metrics(Family family, long rowCount) {
        double height = switch (family) {
            case LINEAR -> LINEAR_HEIGHT_PT;
            case SCRIPT -> SCRIPT_HEIGHT_PT;
            case SCRIPT_FRACTION -> SCRIPT_FRACTION_HEIGHT_PT;
            case SCRIPT_FRACTION_MIXED -> SCRIPT_FRACTION_MIXED_HEIGHT_PT;
            case ORDINARY_FRACTION -> ORDINARY_FRACTION_HEIGHT_PT;
            case NESTED_FRACTION -> NESTED_FRACTION_HEIGHT_PT;
            case TEXT_FRACTION -> TEXT_FRACTION_HEIGHT_PT;
            case SQRT -> SQRT_HEIGHT_PT;
            case SQRT_NESTED -> SQRT_NESTED_HEIGHT_PT;
            case SQRT_FRACTION -> SQRT_FRACTION_HEIGHT_PT;
            case ARRAY -> Math.max(ARRAY_MIN_HEIGHT_PT, Math.max(1L, rowCount) * ARRAY_ROW_HEIGHT_PT);
            case ACCENT -> ACCENT_HEIGHT_PT;
        };
        return new FamilyMetrics(family, height, true);
    }

    private MathTypeStructureMetrics() {
    }
}
