package com.lz.paperword.core.render;

import org.junit.jupiter.api.Test;

import java.io.IOException;
import java.nio.ByteBuffer;
import java.nio.ByteOrder;
import java.util.ArrayList;
import java.util.List;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

class VectorWmfFormulaRendererTest {

    @Test
    void linearFormulaUsesVectorTextRecordsInsteadOfStretchDib() throws IOException {
        byte[] wmf = VectorWmfFormulaRenderer.render("2.25\\div 0.9=2.5", 68.0d, 13.0d);
        List<Integer> records = records(wmf);

        assertTrue(VectorWmfFormulaRenderer.canRender(""));
        assertTrue(records.contains(0x02FB), "vector WMF should create a font");
        assertTrue(records.contains(0x0A32), "vector WMF should draw formula text with ExtTextOut");
        assertFalse(records.contains(0x0F43), "linear vector WMF must not embed a DIB bitmap");
    }

    @Test
    void shortLinearEquationsUseReadableFontHeightWithoutAddingSpaces() throws IOException {
        byte[] equation = VectorWmfFormulaRenderer.render("AE=EF=FB", 66.75d, 12.0d);
        byte[] product = VectorWmfFormulaRenderer.render("12\\times 2=24", 48.0d, 12.0d);
        byte[] label = VectorWmfFormulaRenderer.render("ABC", 30.0d, 12.0d);

        assertTrue(VectorWmfFormulaRenderer.isShortLinearEquation("AE=EF=FB"));
        assertTrue(VectorWmfFormulaRenderer.isShortLinearEquation("12\\times 2=24"));
        assertEquals(1.08d, VectorWmfFormulaRenderer.shortLinearEquationWidthScale("AE=EF=FB"));
        assertEquals(1.08d, VectorWmfFormulaRenderer.shortLinearEquationWidthScale("12\\times 2=24"));
        assertEquals(1.0d, VectorWmfFormulaRenderer.shortLinearEquationWidthScale("AE=EF=FB", 12.75d));
        assertFalse(VectorWmfFormulaRenderer.isShortLinearEquation("ABC"));
        assertFalse(VectorWmfFormulaRenderer.isShortLinearEquation("A+B+C"));
        assertFalse(VectorWmfFormulaRenderer.isShortLinearEquation("12+34"));
        assertFalse(VectorWmfFormulaRenderer.isShortLinearEquation("AB-CD"));
        assertFalse(VectorWmfFormulaRenderer.isShortLinearEquation("S=25+35"));
        assertEquals(1.0d, VectorWmfFormulaRenderer.shortLinearEquationWidthScale("ABC"));
        assertEquals(1.0d, VectorWmfFormulaRenderer.shortLinearEquationWidthScale("A+B+C"));
        assertEquals(1.0d, VectorWmfFormulaRenderer.shortLinearEquationWidthScale("S=25+35"));
        int equationFontHeight = maxSelectedFontHeightTwips(equation);
        int productFontHeight = textFontHeightTwips(product, "12");
        assertTrue(equationFontHeight > 195,
            "short equality formulas should use readable preview font height: " + equationFontHeight);
        assertTrue(productFontHeight > 195,
            "short product formulas should use readable short-linear height: " + productFontHeight);
        assertTrue(totalTextDx(equation) > textDxTotal(label, "ABC"),
            "short-linear calibration should widen glyph advances, not add spaces");
        assertTrue(maxTextRightCoordinate(equation) <= 66.75d * 20.0d);
        assertTrue(maxTextRightCoordinate(product) <= 48.0d * 20.0d);
    }

    @Test
    void limitedStructuredPlaceholdersCanRenderAsVectorText() {
        assertTrue(VectorWmfFormulaRenderer.canRender("x^{\\frac{1}{2}}"));
        assertTrue(VectorWmfFormulaRenderer.canRender("\\sqrt{\\frac{1}{2}}"));
        assertTrue(VectorWmfFormulaRenderer.canRender("{\\sqrt{}}"));
        assertFalse(VectorWmfFormulaRenderer.canRender("\\sqrt[3]{8}"),
            "optional root indexes must not be vectorized until they are emitted in WMF");
        assertTrue(VectorWmfFormulaRenderer.canRender("\\underline{ }"));
    }

    @Test
    void sqrtLayoutsUseStructureMetricHeightFamilies() throws IOException {
        byte[] simpleRoot = VectorWmfFormulaRenderer.render("\\sqrt{2}", 28.0d,
            MathTypeStructureMetrics.SQRT_HEIGHT_PT);
        byte[] fractionRoot = VectorWmfFormulaRenderer.render("\\sqrt{1+\\frac{a}{b}}", 54.0d,
            MathTypeStructureMetrics.SQRT_FRACTION_HEIGHT_PT);
        byte[] rootInFraction = VectorWmfFormulaRenderer.render("\\frac{\\sqrt{a^{2}+b^{2}}}{2}", 54.0d,
            MathTypeStructureMetrics.SQRT_FRACTION_HEIGHT_PT);

        assertFalse(records(simpleRoot).contains(0x0F43));
        assertFalse(records(fractionRoot).contains(0x0F43));
        assertFalse(records(rootInFraction).contains(0x0F43));
        assertTrue(maxTextYCoordinate(simpleRoot) <= MathTypeStructureMetrics.SQRT_HEIGHT_PT * 20.0d);
        assertTrue(maxPolylineYCoordinate(simpleRoot) <= MathTypeStructureMetrics.SQRT_HEIGHT_PT * 20.0d);
        assertTrue(maxTextYCoordinate(fractionRoot) <= MathTypeStructureMetrics.SQRT_FRACTION_HEIGHT_PT * 20.0d);
        assertTrue(maxPolylineYCoordinate(fractionRoot) > MathTypeStructureMetrics.SQRT_HEIGHT_PT * 20.0d,
            "sqrt+fraction radical lines should use the taller MathType sqrt_fraction height family");
        assertTrue(maxPolylineYCoordinate(fractionRoot) <= MathTypeStructureMetrics.SQRT_FRACTION_HEIGHT_PT * 20.0d);
        assertEquals((int) (MathTypeStructureMetrics.SQRT_FRACTION_HEIGHT_PT * 20.0d), windowExtY(rootInFraction));
        assertTrue(maxTextYCoordinate(rootInFraction) > MathTypeStructureMetrics.SQRT_HEIGHT_PT * 20.0d,
            "fractions containing roots should not keep the ordinary 28pt fraction geometry");
        assertTrue(maxTextYCoordinate(rootInFraction) <= MathTypeStructureMetrics.SQRT_FRACTION_HEIGHT_PT * 20.0d);
        assertTrue(maxPolylineYCoordinate(rootInFraction) <= MathTypeStructureMetrics.SQRT_FRACTION_HEIGHT_PT * 20.0d);
        assertFalse(hasCompactRootProfile(simpleRoot),
            "ordinary sqrt roots should not receive the compact sqrt-fraction root profile");
        assertFalse(hasCompactRootProfile(fractionRoot),
            "mixed sqrt bodies should not receive the whole-body fraction root profile");
        assertFalse(hasCompactRootProfile(rootInFraction),
            "roots inside ordinary fractions should not receive the compact root profile");
    }

    @Test
    void sqrtBodyFractionsUseCompactSlotScaleOnlyWhenBodyIsWholeFraction() throws IOException {
        byte[] simpleBodyFraction = VectorWmfFormulaRenderer.render("\\sqrt{\\frac{l}{g}}", 54.0d,
            MathTypeStructureMetrics.SQRT_FRACTION_HEIGHT_PT);
        byte[] standaloneABodyFraction = VectorWmfFormulaRenderer.render("\\sqrt{\\frac{a}{b}}", 54.0d,
            MathTypeStructureMetrics.SQRT_FRACTION_HEIGHT_PT);
        byte[] mixedBodyFraction = VectorWmfFormulaRenderer.render("\\sqrt{1+\\frac{a}{b}}", 54.0d,
            MathTypeStructureMetrics.SQRT_FRACTION_HEIGHT_PT);
        byte[] nestedBodyFraction = VectorWmfFormulaRenderer.render("\\sqrt{\\sqrt{\\frac{a}{b}}}", 72.0d,
            58.0d);

        assertTrue(VectorWmfFormulaRenderer.isWholeSimpleFraction("\\frac{l}{g}"));
        assertTrue(VectorWmfFormulaRenderer.isWholeSimpleFraction("{\\frac{l}{g}}"));
        assertTrue(VectorWmfFormulaRenderer.isWholeSimpleFraction("\\frac{\\pi}{2}"));
        assertFalse(VectorWmfFormulaRenderer.isWholeSimpleFraction("1+\\frac{a}{b}"));
        assertFalse(VectorWmfFormulaRenderer.isWholeSimpleFraction("\\sqrt{\\frac{l}{g}}"));
        assertFalse(VectorWmfFormulaRenderer.isWholeSimpleFraction("\\dfrac{l}{g}"));
        assertFalse(VectorWmfFormulaRenderer.isWholeSimpleFraction("\\frac{1+\\frac{a}{b}}{2+\\frac{c}{d}}"));
        assertFalse(VectorWmfFormulaRenderer.isWholeSimpleFraction("\\frac{\\sqrt{x}}{g}"));
        assertFalse(VectorWmfFormulaRenderer.isWholeSimpleFraction("\\frac{\\text{速度}}{g}"));
        assertTrue(maxTextYCoordinate(simpleBodyFraction) - minTextYCoordinate(simpleBodyFraction)
                < maxTextYCoordinate(mixedBodyFraction) - minTextYCoordinate(mixedBodyFraction),
            "a pure fraction body inside a radical should use a tighter vertical slot");
        assertTrue(maxTextRightCoordinate(simpleBodyFraction) < maxTextRightCoordinate(mixedBodyFraction),
            "a pure fraction body inside a radical should not keep full display fraction advance");
        assertTrue(textDxTotal(nestedBodyFraction, "a") > textDxTotal(standaloneABodyFraction, "a"),
            "pure fraction bodies nested inside another radical should use a more readable nested scale");
        assertTrue(maxTextRightCoordinate(nestedBodyFraction) > maxTextRightCoordinate(simpleBodyFraction),
            "nested sqrt-body fractions should widen relative to standalone compact sqrt-body fractions");
        assertTrue(maxTextRightCoordinate(nestedBodyFraction) >= 22.8d * 20.0d,
            "nested sqrt-body fraction stroke overhangs must not shrink the inner a/b glyph slot");
        List<Polyline> lines = polylines(simpleBodyFraction);
        Polyline radical = lines.stream()
            .filter(VectorWmfFormulaRendererTest::isRadicalPolyline)
            .max((left, right) -> Integer.compare(left.pointCount(), right.pointCount()))
            .orElseThrow();
        Polyline hook = lines.stream()
            .filter(line -> line.pointCount() == 2 && line.x1() < radical.x1() && line.x2() == radical.x1())
            .findFirst()
            .orElseThrow();
        Polyline fractionBar = lines.stream()
            .filter(line -> line.pointCount() == 2 && line.x1() >= radicalTopX(radical) && line.x2() > line.x1())
            .findFirst()
            .orElseThrow();
        assertTrue(hook.x1() < radical.x1() && hook.x2() == radical.x1(),
            "compact radical hook should start left of the main stroke and join it");
        double expectedHookRatio = MathTypeStructureMetrics.SQRT_TALL_COMPACT_HOOK_X_PT
            / MathTypeStructureMetrics.SQRT_TALL_COMPACT_CHECK_TOP_X_PT;
        double actualHookRatio = (double) (radical.x1() - hook.x1())
            / (double) (radicalTopX(radical) - radical.x1());
        assertTrue(Math.abs(actualHookRatio - expectedHookRatio) <= 0.08d,
            "compact radical hook length should stay metric-driven after preview scaling");
        assertCloseTwips(MathTypeStructureMetrics.SQRT_FRACTION_HEIGHT_PT
            * MathTypeStructureMetrics.SQRT_TALL_COMPACT_HOOK_END_Y_RATIO, hook.y2());
        int compactTopIndex = radical.pointCount() - 2;
        int compactTopLeadIndex = compactTopIndex - 1;
        int compactShoulderIndex = compactTopLeadIndex - 1;
        int compactMidIndex = compactShoulderIndex - 1;
        int compactLowerTransitionIndex = compactMidIndex - 1;
        Polyline shadow = lines.stream()
            .filter(line -> line.x1() > radical.x1() && line.pointCount() == 3)
            .filter(line -> line.x2() < radical.x(compactMidIndex))
            .findFirst()
            .orElseThrow();
        Polyline upperProfile = lines.stream()
            .filter(line -> line.pointCount() == 3)
            .filter(line -> line.x1() > radical.x(compactMidIndex))
            .filter(line -> line.x1() < radical.x(compactTopIndex))
            .findFirst()
            .orElseThrow();
        double expectedShadowRatio = MathTypeStructureMetrics.SQRT_TALL_COMPACT_SHADOW_X_OFFSET_PT
            / MathTypeStructureMetrics.SQRT_TALL_COMPACT_CHECK_TOP_X_PT;
        double actualShadowRatio = (double) (shadow.x1() - radical.x1())
            / (double) (radicalTopX(radical) - radical.x1());
        assertTrue(Math.abs(actualShadowRatio - expectedShadowRatio) <= 0.05d,
            "compact radical shadow stroke should thicken only the radical, not the fraction bar");
        int expectedShadowEndX = radical.x(compactLowerTransitionIndex) + shadow.x1() - radical.x1();
        assertTrue(Math.abs(shadow.x(shadow.pointCount() - 1) - expectedShadowEndX) <= 2,
            "compact radical lower profile should stop at the transition point instead of copying the mid/upper turn");
        assertTrue(shadow.x(shadow.pointCount() - 1) < radical.x(compactMidIndex),
            "compact radical lower profile should not duplicate the vertical mid leg");
        assertCloseTwips(MathTypeStructureMetrics.SQRT_TALL_COMPACT_UPPER_PROFILE_X_OFFSET_PT,
            upperProfile.x1() - radical.x(compactShoulderIndex));
        assertCloseTwips(MathTypeStructureMetrics.SQRT_TALL_COMPACT_UPPER_PROFILE_Y_OFFSET_PT,
            upperProfile.y1() - radical.y(compactShoulderIndex));
        assertCloseTwips(MathTypeStructureMetrics.SQRT_TALL_COMPACT_UPPER_PROFILE_X_OFFSET_PT,
            upperProfile.x2() - radical.x(compactTopIndex));
        assertTrue(upperProfile.x1() > radical.x(3) && upperProfile.x2() < radical.x2(),
            "compact upper profile should thicken only the upper turn without extending the top bar");
        assertTrue(upperProfile.x2() <= radical.x(compactTopIndex),
            "compact upper profile must not extend right of the radical top turn");
        assertTrue(upperProfile.y2() >= radical.y(compactTopIndex),
            "compact upper profile must stay inside the radical top edge instead of raising the bbox");
        int radicalTopX = radicalTopX(radical);
        assertTrue(radicalTopX <= fractionBar.x1(),
            "scaled sqrt-body fractions should not protrude left of the radical top turn");
        assertTrue(radicalTopX <= minTextXCoordinate(simpleBodyFraction),
            "scaled sqrt-body fraction glyphs should not protrude left of the radical top turn");
        int fractionGap = fractionBar.x1() - radicalTopX;
        assertTrue(fractionGap >= 8 && fractionGap <= 18,
            "scaled sqrt-body fraction bar should keep a narrow but positive gap after the radical top turn");
        assertTrue(fractionGap <= 45,
            "scaled sqrt-body fractions should sit close to the radical top turn, not float in an oversized slot");
        assertTrue(fractionGap <= Math.round((MathTypeStructureMetrics.SQRT_BODY_LEFT_PAD_PT
                - MathTypeStructureMetrics.SQRT_BODY_FRACTION_LEFT_PAD_PT) * 20.0d) + 2,
            "sqrt-body fractions should use the tighter dedicated body pad instead of the ordinary sqrt body slot");
        assertTrue(radical.x2() >= fractionBar.x2(),
            "scaled sqrt-body fractions must keep the radical top bar covering the fraction bar");
        assertCloseTwips(MathTypeStructureMetrics.SQRT_FRACTION_HEIGHT_PT
            * MathTypeStructureMetrics.SQRT_TALL_COMPACT_CHECK_LOW_Y_RATIO, radical.y(1));
        assertTrue(radical.y(1) < Math.round(MathTypeStructureMetrics.SQRT_FRACTION_HEIGHT_PT
                * MathTypeStructureMetrics.SQRT_TALL_CHECK_LOW_Y_RATIO * 20.0d),
            "compact tall radicals should lift the lower check point without changing the top bar");
        double lowerTransitionRatio = (double) (radical.x(2) - radical.x1())
            / (double) (radicalTopX(radical) - radical.x1());
        double expectedLowerTransitionRatio =
            MathTypeStructureMetrics.SQRT_TALL_COMPACT_CHECK_LOWER_TRANSITION_X_PT
                / MathTypeStructureMetrics.SQRT_TALL_COMPACT_CHECK_TOP_X_PT;
        assertTrue(Math.abs(lowerTransitionRatio - expectedLowerTransitionRatio) <= 0.05d,
            "compact tall radicals should use a metric-driven lower transition point");
        assertCloseTwips(MathTypeStructureMetrics.SQRT_FRACTION_HEIGHT_PT
            * MathTypeStructureMetrics.SQRT_TALL_COMPACT_CHECK_LOWER_TRANSITION_Y_RATIO, radical.y(2));
        assertTrue(radical.x(1) < radical.x(2) && radical.x(2) < radical.x(3),
            "compact lower transition point should stay between the low point and mid point");
        assertTrue(radical.y(2) < radical.y(1) && radical.y(2) < radical.y(3),
            "compact lower transition point should soften the lower leg inside the existing bbox");
        double compactShoulderRatio = (double) (radical.x(compactShoulderIndex) - radical.x1())
            / (double) (radical.x(compactTopIndex) - radical.x1());
        double expectedCompactShoulderRatio = MathTypeStructureMetrics.SQRT_TALL_COMPACT_CHECK_SHOULDER_X_PT
            / MathTypeStructureMetrics.SQRT_TALL_COMPACT_CHECK_TOP_X_PT;
        assertTrue(Math.abs(compactShoulderRatio - expectedCompactShoulderRatio) <= 0.08d,
            "compact tall radicals should use the opened shoulder profile, not the old near-vertical leg");
        assertCloseTwips(MathTypeStructureMetrics.SQRT_FRACTION_HEIGHT_PT
            * MathTypeStructureMetrics.SQRT_TALL_COMPACT_CHECK_SHOULDER_Y_RATIO,
            radical.y(compactShoulderIndex));
        double compactTopLeadRatio = (double) (radical.x(compactTopLeadIndex) - radical.x1())
            / (double) (radical.x(compactTopIndex) - radical.x1());
        double expectedCompactTopLeadRatio = MathTypeStructureMetrics.SQRT_TALL_COMPACT_CHECK_TOP_LEAD_X_PT
            / MathTypeStructureMetrics.SQRT_TALL_COMPACT_CHECK_TOP_X_PT;
        assertTrue(Math.abs(compactTopLeadRatio - expectedCompactTopLeadRatio) <= 0.05d,
            "compact top-lead point should stay metric-driven after preview scaling");
        assertCloseTwips(MathTypeStructureMetrics.SQRT_FRACTION_HEIGHT_PT
            * MathTypeStructureMetrics.SQRT_TALL_COMPACT_CHECK_TOP_LEAD_Y_RATIO,
            radical.y(compactTopLeadIndex));
        assertTrue(radical.x(compactShoulderIndex) < radical.x(compactTopLeadIndex)
                && radical.x(compactTopLeadIndex) < radical.x(compactTopIndex),
            "compact tall radicals should add an internal top-lead point without moving the top turn");
        int sqrtBarVisualInset = fractionBar.x1() - minTextXCoordinate(simpleBodyFraction);
        assertTrue(sqrtBarVisualInset <= 12,
            "sqrt-body fraction bars should visually track the glyph width instead of looking like a short dash");
        List<Polyline> nestedBodyFractionLines = polylines(nestedBodyFraction);
        Polyline innerNestedRoot = sortedRadicals(nestedBodyFractionLines).get(1);
        Polyline nestedFractionBar = nestedBodyFractionLines.stream()
            .filter(line -> line.pointCount() == 2)
            .findFirst()
            .orElseThrow();
        assertTrue(radicalTopX(innerNestedRoot) <= nestedFractionBar.x1(),
            "widened nested sqrt-body fraction bars should not protrude left of the inner radical top turn");
        assertTrue(innerNestedRoot.x2() >= nestedFractionBar.x2(),
            "widened nested sqrt-body fractions must keep the inner radical top bar covering the fraction bar");
        assertTrue(innerNestedRoot.x2() <= 1530,
            "nested compact sqrt tuning must not extend the inner top bar beyond the v207 stable bbox");
        int nestedFractionGap = nestedFractionBar.x1() - radicalTopX(innerNestedRoot);
        assertTrue(nestedFractionGap >= fractionGap
                + Math.round(MathTypeStructureMetrics.SQRT_NESTED_BODY_FRACTION_LEFT_EXTRA_PT * 20.0d) - 2,
            "nested sqrt-body fractions should get dedicated inner clearance from the radical turn");
        assertTrue(textFontHeightTwips(mixedBodyFraction, "a") >= 220,
            "mixed root bodies such as 1+frac must not receive the pure-body fraction shrink");
        assertEquals((int) (MathTypeStructureMetrics.SQRT_FRACTION_HEIGHT_PT * 20.0d),
            windowExtY(simpleBodyFraction));
    }

    @Test
    void sqrtRootShapeCoordinatesComeFromStructureMetrics() throws IOException {
        byte[] simpleRoot = VectorWmfFormulaRenderer.render("\\sqrt{2}", 28.0d,
            MathTypeStructureMetrics.SQRT_HEIGHT_PT);
        byte[] fractionRoot = VectorWmfFormulaRenderer.render("\\sqrt{1+\\frac{a}{b}}", 54.0d,
            MathTypeStructureMetrics.SQRT_FRACTION_HEIGHT_PT);
        byte[] rootInFraction = VectorWmfFormulaRenderer.render("\\frac{\\sqrt{a^{2}+b^{2}}}{2}", 54.0d,
            MathTypeStructureMetrics.SQRT_FRACTION_HEIGHT_PT);
        byte[] nestedOnlyRoot = VectorWmfFormulaRenderer.render("\\sqrt{\\sqrt{x}}", 54.0d, 32.0d);
        byte[] nestedRoot = VectorWmfFormulaRenderer.render("\\sqrt{1+\\sqrt{x}}", 64.0d, 32.0d);
        byte[] nestedFractionRoot = VectorWmfFormulaRenderer.render("\\sqrt{1+\\sqrt{\\frac{a}{b}}}", 92.0d,
            58.0d);
        List<Polyline> lines = polylines(simpleRoot);
        List<Polyline> fractionLines = polylines(fractionRoot);
        List<Polyline> rootInFractionLines = polylines(rootInFraction);
        List<Polyline> nestedOnlyLines = polylines(nestedOnlyRoot);
        List<Polyline> nestedLines = polylines(nestedRoot);
        List<Polyline> nestedFractionLines = polylines(nestedFractionRoot);

        assertEquals(1, lines.size());
        assertEquals(2, fractionLines.size());
        assertEquals(2, rootInFractionLines.size());
        assertEquals(2, nestedOnlyLines.size());
        assertEquals(2, nestedLines.size());
        assertTrue(nestedFractionLines.size() >= 3,
            "nested fraction roots may include compact radical hook strokes in addition to required bars");
        assertTrue(nestedLines.stream().allMatch(VectorWmfFormulaRendererTest::isRadicalPolyline),
            "nested root radicals must keep radical polylines after layout translation");
        assertTrue(radicalTopGapTwips(nestedLines) >= Math.round(
                MathTypeStructureMetrics.SQRT_NESTED_BODY_Y_EXTRA_PT * 20.0d),
            "nested root top bars should have visible vertical separation");
        List<Polyline> nestedOnlyRadicals = sortedRadicals(nestedOnlyLines);
        int nestedRootHorizontalGap = radicalTopX(nestedOnlyRadicals.get(1))
            - radicalTopX(nestedOnlyRadicals.get(0));
        assertTrue(nestedRootHorizontalGap >= Math.round((MathTypeStructureMetrics.SQRT_BODY_LEFT_PAD_PT
                + MathTypeStructureMetrics.SQRT_NESTED_BODY_LEFT_EXTRA_PT) * 20.0d) - 2,
            "nested root bodies should include ordinary body pad plus dedicated horizontal breathing room");
        assertTrue(radicalTopGapTwips(nestedFractionLines) >= Math.round(
                MathTypeStructureMetrics.SQRT_NESTED_BODY_Y_EXTRA_PT * 20.0d),
            "nested fraction root top bars should have visible vertical separation");
        List<Polyline> nestedFractionRadicals = sortedMainRadicals(nestedFractionLines);
        assertEquals(2, nestedFractionRadicals.size());
        int nestedFractionHorizontalGap = radicalTopX(nestedFractionRadicals.get(1))
            - radicalTopX(nestedFractionRadicals.get(0));
        assertTrue(nestedFractionHorizontalGap >= Math.round((MathTypeStructureMetrics.SQRT_BODY_LEFT_PAD_PT
                + MathTypeStructureMetrics.SQRT_NESTED_BODY_LEFT_EXTRA_PT) * 20.0d) - 2,
            "nested fraction roots should use the same dedicated horizontal breathing room as simple nested roots");
        Polyline root = lines.get(0);
        Polyline tallRoot = fractionLines.stream()
            .filter(VectorWmfFormulaRendererTest::isRadicalPolyline)
            .max((left, right) -> Integer.compare(left.y(1), right.y(1)))
            .orElseThrow();
        Polyline scaledFractionRoot = rootInFractionLines.stream()
            .filter(VectorWmfFormulaRendererTest::isRadicalPolyline)
            .findFirst()
            .orElseThrow();
        List<Polyline> tallNestedFractionRoots = nestedFractionLines.stream()
            .filter(VectorWmfFormulaRendererTest::isMainRadicalPolyline)
            .toList();
        assertEquals(4, root.pointCount());
        assertTrue(tallRoot.pointCount() > 4,
            "tall sqrt radicals should use a richer multi-segment stroke");
        int rootX = root.x1();
        assertCloseTwips(MathTypeStructureMetrics.SQRT_HEIGHT_PT
            * MathTypeStructureMetrics.SQRT_LEFT_DESCENT_RATIO, root.y1());
        int midX = root.x(1) - rootX;
        int topX = root.x(2) - rootX;
        double expectedMidRatio = MathTypeStructureMetrics.SQRT_CHECK_MID_X_PT
            / MathTypeStructureMetrics.SQRT_CHECK_TOP_X_PT;
        assertTrue(Math.abs(((double) midX / (double) topX) - expectedMidRatio) <= 0.05d,
            "sqrt checkmark x ratio should stay metric-driven");
        int tallTopIndex = tallRoot.pointCount() - 2;
        int tallShoulderIndex = tallTopIndex - 1;
        int tallMidIndex = tallShoulderIndex - 1;
        int tallTopX = tallRoot.x(tallTopIndex) - tallRoot.x1();
        int tallShoulderX = tallRoot.x(tallShoulderIndex) - tallRoot.x1();
        int tallMidX = tallRoot.x(tallMidIndex) - tallRoot.x1();
        int tallLowX = tallRoot.x(1) - tallRoot.x1();
        double expectedTallTopRatio = MathTypeStructureMetrics.SQRT_TALL_CHECK_TOP_X_PT
            / MathTypeStructureMetrics.SQRT_CHECK_TOP_X_PT;
        assertTrue(Math.abs(((double) tallTopX / (double) topX) - expectedTallTopRatio) <= 0.08d,
            "tall sqrt checkmark top turn should be narrower than the ordinary sqrt turn");
        assertTrue(tallTopX >= Math.round(topX * 0.88d),
            "tall sqrt checkmark should stay open enough to avoid an almost vertical left leg");
        double expectedTallMidRatio = MathTypeStructureMetrics.SQRT_TALL_CHECK_MID_X_PT
            / MathTypeStructureMetrics.SQRT_TALL_CHECK_TOP_X_PT;
        assertTrue(Math.abs(((double) tallMidX / (double) tallTopX) - expectedTallMidRatio) <= 0.08d,
            "tall sqrt checkmark midpoint should stay metric-driven after WMF scaling");
        double expectedTallShoulderRatio = MathTypeStructureMetrics.SQRT_TALL_CHECK_SHOULDER_X_PT
            / MathTypeStructureMetrics.SQRT_TALL_CHECK_TOP_X_PT;
        assertTrue(Math.abs(((double) tallShoulderX / (double) tallTopX) - expectedTallShoulderRatio) <= 0.08d,
            "tall sqrt shoulder should stay metric-driven after WMF scaling");
        assertTrue(tallLowX < tallMidX && tallRoot.y(1) > tallRoot.y(tallMidIndex),
            "tall sqrt should add a low shoulder before the rising stroke");
        assertTrue(tallMidX < tallShoulderX && tallShoulderX < tallTopX
                && tallRoot.y(tallMidIndex) > tallRoot.y(tallShoulderIndex)
                && tallRoot.y(tallShoulderIndex) > tallRoot.y(tallTopIndex),
            "tall sqrt should add a middle shoulder that smooths the rising stroke before the top turn");
        assertTrue(((double) tallShoulderX / (double) tallTopX) >= expectedTallShoulderRatio - 0.08d,
            "tall sqrt shoulder should stay open near the top turn instead of collapsing into a vertical stroke");
        assertTrue(((double) tallMidX / (double) tallTopX) < ((double) midX / (double) topX),
            "tall sqrt checkmark midpoint should open the rising stroke instead of drawing an almost vertical leg");
        assertTrue(minTextXCoordinate(fractionRoot) - tallRoot.x1()
            >= tallTopX,
            "tall sqrt bodies should remain to the right of the narrowed radical top turn");
        assertCloseTwips(MathTypeStructureMetrics.SQRT_HEIGHT_PT
            - MathTypeStructureMetrics.SQRT_BOTTOM_PAD_PT, root.y(1));
        assertCloseTwips(MathTypeStructureMetrics.SQRT_FRACTION_HEIGHT_PT
            - MathTypeStructureMetrics.SQRT_TALL_BOTTOM_PAD_PT, tallRoot.y(tallMidIndex));
        assertTrue(tallRoot.y(tallMidIndex) < (MathTypeStructureMetrics.SQRT_FRACTION_HEIGHT_PT
                - MathTypeStructureMetrics.SQRT_BOTTOM_PAD_PT) * 20.0d,
            "tall roots should lift the checkmark turn instead of drawing an overlong left descender");
        assertTrue(isRadicalPolyline(scaledFractionRoot),
            "root inside a fraction should still keep a radical polyline after scaled placement");
        assertTrue(tallNestedFractionRoots.stream().allMatch(VectorWmfFormulaRendererTest::hasLiftedTallRootTurn),
            "nested fraction roots should also use the lifted tall-root checkmark turn");
        assertTrue(tallNestedFractionRoots.stream().allMatch(line -> line.x2() >= maxTextRightCoordinate(nestedFractionRoot)),
            "nested fraction root top bars should keep covering the nested body text");
        assertCloseTwips(MathTypeStructureMetrics.SQRT_TOP_Y_PT, root.y(2));
        assertCloseTwips(MathTypeStructureMetrics.SQRT_TOP_Y_PT, root.y(3));
        assertTrue(minTextXCoordinate(simpleRoot) - rootX >= topX);
        assertTrue(minTextYCoordinate(simpleRoot) >= (int) Math.round(MathTypeStructureMetrics.SQRT_BODY_Y_OFFSET_PT
            * 20.0d));
    }

    @Test
    void tallSqrtThresholdUsesSharedStructureMetric() throws IOException {
        double ordinaryHeight = MathTypeStructureMetrics.SQRT_HEIGHT_PT
            + MathTypeStructureMetrics.SQRT_TALL_EXTRA_HEIGHT_PT;
        double tallHeight = ordinaryHeight + 0.05d;

        assertFalse(MathTypeStructureMetrics.isTallSqrt(ordinaryHeight));
        assertTrue(MathTypeStructureMetrics.isTallSqrt(tallHeight));
        assertEquals(MathTypeStructureMetrics.SQRT_BOTTOM_PAD_PT,
            MathTypeStructureMetrics.sqrtBottomPadPt(ordinaryHeight));
        assertEquals(MathTypeStructureMetrics.SQRT_TALL_BOTTOM_PAD_PT,
            MathTypeStructureMetrics.sqrtBottomPadPt(tallHeight));

        byte[] ordinaryWmf = VectorWmfFormulaRenderer.render("\\sqrt{x}", 34.0d, ordinaryHeight);
        byte[] tallWmf = VectorWmfFormulaRenderer.render("\\sqrt{1+\\frac{a}{b}}", 54.0d, tallHeight);
        Polyline ordinaryRoot = polylines(ordinaryWmf).stream()
            .filter(VectorWmfFormulaRendererTest::isRadicalPolyline)
            .findFirst()
            .orElseThrow();
        Polyline tallRoot = polylines(tallWmf).stream()
            .filter(VectorWmfFormulaRendererTest::isRadicalPolyline)
            .findFirst()
            .orElseThrow();

        assertEquals(4, ordinaryRoot.pointCount(),
            "simple roots should still use the ordinary radical stroke");
        assertTrue(tallRoot.pointCount() > 4,
            "fraction roots should switch to the tall radical stroke through the shared tall-sqrt predicate");
    }

    @Test
    void nestedSqrtOffsetRequiresStructuralSqrtCommand() throws IOException {
        byte[] nestedSqrt = VectorWmfFormulaRenderer.render("\\sqrt{1+\\sqrt{x}}", 44.0d,
            MathTypeStructureMetrics.SQRT_HEIGHT_PT);

        assertFalse(VectorWmfFormulaRenderer.hasSqrtCommandOutsideText("\\text{\\sqrt}"));
        assertFalse(VectorWmfFormulaRenderer.hasSqrtCommandOutsideText("\\mathrm{\\sqrt}+x"));
        assertTrue(VectorWmfFormulaRenderer.hasSqrtCommandOutsideText("1+\\sqrt{x}"));
        assertEquals(0, VectorWmfFormulaRenderer.sqrtCommandDepthOutsideText("\\text{\\sqrt}"));
        assertEquals(1, VectorWmfFormulaRenderer.sqrtCommandDepthOutsideText("1+\\sqrt{x}"));
        assertEquals(2, VectorWmfFormulaRenderer.sqrtCommandDepthOutsideText("x+\\sqrt{y+\\sqrt{z}}"));
        assertEquals(2, polylines(nestedSqrt).size());
        assertTrue(radicalTopGapTwips(polylines(nestedSqrt)) >= Math.round(
                MathTypeStructureMetrics.SQRT_NESTED_BODY_Y_EXTRA_PT * 20.0d),
            "structural nested sqrt should keep the extra vertical separation");
    }

    @Test
    void deepNestedSqrtStaysInsideProductionSqrtEnvelope() throws IOException {
        byte[] deepNested = VectorWmfFormulaRenderer.render("\\sqrt{x+\\sqrt{y+\\sqrt{z}}}", 62.0d,
            MathTypeStructureMetrics.SQRT_NESTED_HEIGHT_PT);
        List<Polyline> lines = polylines(deepNested);

        assertEquals((int) (MathTypeStructureMetrics.SQRT_NESTED_HEIGHT_PT * 20.0d), windowExtY(deepNested));
        assertEquals(3, lines.stream().filter(VectorWmfFormulaRendererTest::isRadicalPolyline).count());
        assertTrue(maxPolylineYCoordinate(deepNested) <= MathTypeStructureMetrics.SQRT_NESTED_HEIGHT_PT * 20.0d);
        assertTrue(maxTextYCoordinate(deepNested) <= MathTypeStructureMetrics.SQRT_NESTED_HEIGHT_PT * 20.0d);
        List<Polyline> radicals = sortedRadicals(lines);
        int outerToMiddleGap = radicalTopY(radicals.get(1)) - radicalTopY(radicals.get(0));
        int middleToInnerGap = radicalTopY(radicals.get(2)) - radicalTopY(radicals.get(1));
        assertTrue(outerToMiddleGap > Math.round(MathTypeStructureMetrics.SQRT_NESTED_BODY_Y_EXTRA_PT * 20.0d),
            "outer deep-nested root should add depth-scaled top-bar separation");
        assertTrue(middleToInnerGap >= Math.round(MathTypeStructureMetrics.SQRT_NESTED_BODY_Y_EXTRA_PT * 20.0d),
            "inner deep-nested roots should keep at least one nested spacing step");
    }

    @Test
    void scaledStructuredLayoutsPreserveRootPolylinePoints() throws IOException {
        byte[] rootInFraction = VectorWmfFormulaRenderer.render("\\frac{\\sqrt{x}}{2}", 54.0d,
            MathTypeStructureMetrics.SQRT_FRACTION_HEIGHT_PT);

        assertTrue(polylines(rootInFraction).stream().anyMatch(VectorWmfFormulaRendererTest::isRadicalPolyline),
            "root inside a scaled fraction slot must keep the radical shape");
    }

    @Test
    void unifiedBoxLayoutCoversCoreFormulaShapes() throws IOException {
        String linear = "S_{\\Delta AOB}:S_{\\Delta COD}=a^{2}:b^{2}=4:9";
        String nested = "\\frac{1+\\frac{a}{b}}{2+\\frac{c}{d}}";
        String textHeavyFraction = "\\frac{三角形ABD的面积}{三角形CBD的面积}=\\frac{AO}{CO}";
        String complexRoot = "\\sqrt{1+\\frac{a_{1}^{2}}{\\frac{3}{5}}}";

        byte[] linearWmf = VectorWmfFormulaRenderer.render(linear, 170.0d, 18.0d);
        byte[] nestedWmf = VectorWmfFormulaRenderer.render(nested, 82.0d, 53.0d);
        byte[] textHeavyWmf = VectorWmfFormulaRenderer.render(textHeavyFraction, 160.0d,
            MathTypeStructureMetrics.TEXT_FRACTION_HEIGHT_PT);
        byte[] complexRootWmf = VectorWmfFormulaRenderer.render(complexRoot, 82.0d, 62.0d);

        assertTrue(records(linearWmf).contains(0x0A32));
        assertTrue(records(nestedWmf).contains(0x0325));
        assertTrue(records(nestedWmf).contains(0x0A32));
        assertTrue(records(textHeavyWmf).contains(0x0A32));
        assertTrue(records(complexRootWmf).contains(0x0325));
        assertTrue(records(complexRootWmf).contains(0x0A32));
        assertFalse(records(linearWmf).contains(0x0F43));
        assertFalse(records(nestedWmf).contains(0x0F43));
        assertFalse(records(textHeavyWmf).contains(0x0F43));
        assertFalse(records(complexRootWmf).contains(0x0F43));
        assertEquals((int) (MathTypeStructureMetrics.TEXT_FRACTION_HEIGHT_PT * 20.0d),
            windowExtY(textHeavyWmf));
        assertTrue(maxTextYCoordinate(textHeavyWmf) > 18.0d * 20.0d,
            "text-heavy fractions should not be compressed into the old 18pt inline profile");
        assertTrue(maxTextYCoordinate(textHeavyWmf) <= MathTypeStructureMetrics.TEXT_FRACTION_HEIGHT_PT * 20.0d);
        assertTrue(maxTextYCoordinate(nestedWmf) > 30.0d * 20.0d,
            "nested fractions should use the tall MathType height family, not compact unified-box spacing");
    }

    @Test
    void textHeavyFractionsUseSourceHeightFamily() throws IOException {
        String standalone = "\\frac{三角形ABD的面积}{三角形CBD的面积}";
        String mixed = "\\frac{三角形ABD的面积}{三角形CBD的面积}=\\frac{AO}{CO}";

        byte[] standaloneWmf = VectorWmfFormulaRenderer.render(standalone, 112.0d,
            MathTypeStructureMetrics.TEXT_FRACTION_HEIGHT_PT);
        byte[] mixedWmf = VectorWmfFormulaRenderer.render(mixed, 160.0d,
            MathTypeStructureMetrics.TEXT_FRACTION_HEIGHT_PT);

        assertEquals((int) (MathTypeStructureMetrics.TEXT_FRACTION_HEIGHT_PT * 20.0d),
            windowExtY(standaloneWmf));
        assertEquals((int) (MathTypeStructureMetrics.TEXT_FRACTION_HEIGHT_PT * 20.0d), windowExtY(mixedWmf));
        List<Polyline> mixedBars = polylines(mixedWmf);
        assertEquals(2, mixedBars.size());
        Polyline mainBar = mixedBars.stream()
            .filter(line -> line.x2() - line.x1() > 1800)
            .findFirst()
            .orElseThrow();
        Polyline trailingBar = mixedBars.stream()
            .filter(line -> line.x2() - line.x1() < 700)
            .findFirst()
            .orElseThrow();
        int mainNumeratorBaseline = minTextYCoordinateBetween(mixedWmf, mainBar.x1(), mainBar.x2());
        int mainDenominatorBaseline = maxTextYCoordinateBetween(mixedWmf, mainBar.x1(), mainBar.x2());
        int trailingNumeratorBaseline = minTextYCoordinateBetween(mixedWmf, trailingBar.x1(), trailingBar.x2());
        int trailingDenominatorBaseline = maxTextYCoordinateBetween(mixedWmf, trailingBar.x1(), trailingBar.x2());
        int mainNumeratorGap = mainBar.y1() - mainNumeratorBaseline;
        int mainDenominatorGap = mainDenominatorBaseline - mainBar.y1();
        int trailingNumeratorGap = trailingBar.y1() - trailingNumeratorBaseline;
        int trailingDenominatorGap = trailingDenominatorBaseline - trailingBar.y1();
        assertTrue(mainBar.x1() <= minTextXCoordinateBetween(mixedWmf, mainBar.x1(), mainBar.x2()) + 6,
            "text-heavy fraction bar should start close to the text ink, not after a large inset");
        assertTrue(mainBar.x2() >= maxTextRightCoordinateBetween(mixedWmf, mainBar.x1(), mainBar.x2()) - 6,
            "text-heavy fraction bar should cover the wide CJK numerator/denominator text");
        assertTrue(mainNumeratorGap >= 110,
            "text-heavy fraction bar should sit visibly below the numerator baseline");
        assertTrue(mainDenominatorGap >= 110,
            "text-heavy fraction bar should sit visibly above the denominator baseline");
        assertTrue(trailingNumeratorGap < mainNumeratorGap,
            "ordinary trailing fraction in a text-heavy formula should keep ordinary fraction geometry");
        assertTrue(trailingDenominatorGap < mainDenominatorGap,
            "ordinary trailing fraction in a text-heavy formula should not use tall text-heavy denominator spacing");
        assertTrue(records(standaloneWmf).contains(0x0325));
        assertTrue(records(mixedWmf).contains(0x0325));
        assertFalse(records(standaloneWmf).contains(0x0F43));
        assertFalse(records(mixedWmf).contains(0x0F43));
        assertTrue(maxTextYCoordinate(standaloneWmf) > 18.0d * 20.0d);
        assertTrue(maxTextYCoordinate(mixedWmf) > 18.0d * 20.0d,
            "text-heavy mixed fractions must not route through compact inline fraction geometry");
        assertTrue(maxTextYCoordinate(standaloneWmf) <= MathTypeStructureMetrics.TEXT_FRACTION_HEIGHT_PT * 20.0d);
        assertTrue(maxTextYCoordinate(mixedWmf) <= MathTypeStructureMetrics.TEXT_FRACTION_HEIGHT_PT * 20.0d);
        assertTrue(maxPolylineYCoordinate(standaloneWmf) <= MathTypeStructureMetrics.TEXT_FRACTION_HEIGHT_PT * 20.0d);
        assertTrue(maxPolylineYCoordinate(mixedWmf) <= MathTypeStructureMetrics.TEXT_FRACTION_HEIGHT_PT * 20.0d);
    }

    @Test
    void standardGlyphModelKeepsScriptsAndInlineFractionsOnOneSizeLadder() throws IOException {
        String scriptRelation = "S_{\\Delta AOB}:S_{\\Delta COD}=a^{2}:b^{2}=4:9";
        String shortScript = "a_{1}^{2}+b_{2}^{3}=c^{2}";
        String inlineFraction = "AO=1,\\ CO=\\frac{5}{3}";
        String complexRoot = "\\sqrt{1+\\frac{a_{1}^{2}}{\\frac{3}{5}}}";

        byte[] relationWmf = VectorWmfFormulaRenderer.render(scriptRelation, 170.0d, 18.0d);
        byte[] shortScriptWmf = VectorWmfFormulaRenderer.render(shortScript, 85.0d, 18.0d);
        byte[] fractionWmf = VectorWmfFormulaRenderer.render(inlineFraction, 70.0d, 26.0d);

        assertTrue(VectorWmfFormulaRenderer.standardGlyphModel(scriptRelation));
        assertTrue(VectorWmfFormulaRenderer.standardGlyphModel(shortScript));
        assertTrue(VectorWmfFormulaRenderer.standardGlyphModel(inlineFraction));
        assertFalse(VectorWmfFormulaRenderer.standardGlyphModel("\\dfrac{a_{1}}{b}"));
        assertFalse(VectorWmfFormulaRenderer.standardGlyphModel("\\frac{a_{1}}{\\frac{3}{5}}"));
        assertFalse(VectorWmfFormulaRenderer.standardGlyphModel(complexRoot));
        assertFalse(VectorWmfFormulaRenderer.standardGlyphModel(
            "\\frac{三角形ABD的面积}{三角形CBD的面积}=\\frac{AO}{CO}"));
        assertTrue(textFontHeightTwips(relationWmf, "S") >= 185);
        assertTrue(textFontHeightTwips(relationWmf, "AOB") >= 135,
            "script glyphs should not be double-shrunk below the standard ratio");
        assertTrue(textFontHeightTwips(shortScriptWmf, "1") >= 135);
        assertTrue(textFontHeightTwips(fractionWmf, "AO") >= 200);
        assertTrue(textFontHeightTwips(fractionWmf, "5") >= 150);
    }

    @Test
    void standardGlyphAdvanceDoesNotMultiplyVerticalScaleIntoWidth() throws IOException {
        String scriptRelation = "S_{\\Delta AOB}:S_{\\Delta COD}=a^{2}:b^{2}=4:9";

        byte[] normalHeight = VectorWmfFormulaRenderer.render(scriptRelation, 170.0d, 18.0d);
        byte[] tallBox = VectorWmfFormulaRenderer.render(scriptRelation, 170.0d, 27.0d);

        assertTrue(VectorWmfFormulaRenderer.standardGlyphModel(scriptRelation));
        assertEquals(textFontHeightTwips(normalHeight, "S"), textFontHeightTwips(tallBox, "S"));
        assertFalse(records(normalHeight).contains(0x0F43));
        assertFalse(records(tallBox).contains(0x0F43));
        assertTrue(maxTextRightCoordinate(tallBox) < maxTextRightCoordinate(normalHeight) * 1.25d,
            "standard glyph advance should not multiply y scale into width");
    }

    @Test
    void scriptFractionLayoutsPreserveChildWidthScaleWhenAppended() throws IOException {
        byte[] scriptFraction = VectorWmfFormulaRenderer.render("x^{\\frac{a_{1}^{2}}{2}}", 82.0d,
            MathTypeStructureMetrics.SCRIPT_FRACTION_HEIGHT_PT);
        byte[] pureScriptFractions = VectorWmfFormulaRenderer.render(
            "x^{\\frac{1}{2}}+a^{\\frac{2}{3}}", 82.0d,
            MathTypeStructureMetrics.SCRIPT_FRACTION_HEIGHT_PT);
        byte[] mixedScriptFraction = VectorWmfFormulaRenderer.render(
            "S_{\\frac{1}{4}\\mathrm{圆}}=\\frac{1}{4}\\pi r^{2}", 92.0d,
            MathTypeStructureMetrics.SCRIPT_FRACTION_MIXED_HEIGHT_PT);
        byte[] nestedWithScriptFraction = VectorWmfFormulaRenderer.render("\\frac{x^{\\frac{1}{2}}}{2}", 48.0d,
            MathTypeStructureMetrics.NESTED_FRACTION_HEIGHT_PT);
        byte[] standalone = VectorWmfFormulaRenderer.render("a_{1}^{2}", 38.0d, 13.0d);

        assertTrue(VectorWmfFormulaRenderer.canRender("\\frac{x^{\\frac{1}{2}}}{2}"));
        assertFalse(records(scriptFraction).contains(0x0325),
            "script-slot fractions should stay on the slash/unified path, not display stacked bars");
        assertFalse(records(pureScriptFractions).contains(0x0325));
        assertTrue(records(mixedScriptFraction).contains(0x0325),
            "mixed script_fraction formulas still contain a top-level ordinary fraction");
        assertTrue(records(nestedWithScriptFraction).contains(0x0325),
            "script-slot fractions inside a display fraction should stay on the nested display-fraction path");
        assertFalse(records(scriptFraction).contains(0x0F43));
        assertFalse(records(pureScriptFractions).contains(0x0F43));
        assertFalse(records(mixedScriptFraction).contains(0x0F43));
        assertEquals((int) (MathTypeStructureMetrics.SCRIPT_FRACTION_HEIGHT_PT * 20.0d),
            windowExtY(scriptFraction));
        assertEquals((int) (MathTypeStructureMetrics.SCRIPT_FRACTION_HEIGHT_PT * 20.0d),
            windowExtY(pureScriptFractions));
        assertEquals((int) (MathTypeStructureMetrics.SCRIPT_FRACTION_MIXED_HEIGHT_PT * 20.0d),
            windowExtY(mixedScriptFraction));
        assertEquals((int) (MathTypeStructureMetrics.NESTED_FRACTION_HEIGHT_PT * 20.0d),
            windowExtY(nestedWithScriptFraction));
        assertTrue(textDxTotal(scriptFraction, "a") < textDxTotal(standalone, "a"),
            "script fraction child script run should keep widthScale when appended");
        assertTrue(maxTextYCoordinate(scriptFraction) <= MathTypeStructureMetrics.SCRIPT_FRACTION_HEIGHT_PT * 20.0d);
        assertTrue(maxTextYCoordinate(pureScriptFractions)
            <= MathTypeStructureMetrics.SCRIPT_FRACTION_HEIGHT_PT * 20.0d);
        assertTrue(maxTextYCoordinate(mixedScriptFraction)
            <= MathTypeStructureMetrics.SCRIPT_FRACTION_MIXED_HEIGHT_PT * 20.0d);
        assertTrue(maxPolylineYCoordinate(nestedWithScriptFraction)
            >= MathTypeStructureMetrics.NESTED_FRACTION_ABOVE_PT * 20.0d,
            "top-level display fractions containing script-slot fractions must not be recast as script_fraction_mixed");
        assertTrue(maxTextRightCoordinate(scriptFraction) <= 82.0d * 20.0d);
    }

    @Test
    void nestedFractionsStayOnTallDisplayFractionPath() throws IOException {
        byte[] nested = VectorWmfFormulaRenderer.render("\\frac{1+a_{1}^{2}}{2+\\frac{3}{5}}", 82.0d,
            MathTypeStructureMetrics.NESTED_FRACTION_HEIGHT_PT);

        assertTrue(records(nested).contains(0x0325));
        assertFalse(records(nested).contains(0x0F43));
        assertEquals((int) (MathTypeStructureMetrics.NESTED_FRACTION_HEIGHT_PT * 20.0d), windowExtY(nested));
        assertTrue(maxPolylineYCoordinate(nested) >= MathTypeStructureMetrics.NESTED_FRACTION_ABOVE_PT * 20.0d,
            "outer nested fraction bar should use the source-seeded nested fraction split");
        assertTrue(maxTextYCoordinate(nested) > 40.0d * 20.0d,
            "nested display fractions should distribute denominator content through the tall source family");
        assertTrue(maxTextYCoordinate(nested) <= MathTypeStructureMetrics.NESTED_FRACTION_HEIGHT_PT * 20.0d);
        assertTrue(maxTextRightCoordinate(nested) <= 82.0d * 20.0d);
    }

    @Test
    void displayFractionCommandsUseVectorFractionGeometry() throws IOException {
        byte[] dfrac = VectorWmfFormulaRenderer.render("\\dfrac{1}{2}", 18.0d, 28.0d);
        byte[] cfrac = VectorWmfFormulaRenderer.render("\\cfrac{a}{b}", 22.0d, 28.0d);

        assertTrue(records(dfrac).contains(0x0325));
        assertTrue(records(cfrac).contains(0x0325));
        assertFalse(records(dfrac).contains(0x0F43));
        assertFalse(records(cfrac).contains(0x0F43));
        assertTrue(maxTextYCoordinate(dfrac) > 18.0d * 20.0d,
            "display fraction commands should not be squeezed into a linear 13pt profile");
    }

    @Test
    void fractionScannerRequiresLatexCommandBoundary() {
        assertFalse(VectorWmfFormulaRenderer.canRender("\\fraction{x}"),
            "command prefixes such as \\fraction must not be parsed as \\frac");
        assertFalse(VectorWmfFormulaRenderer.canRender("\\dfraction{x}"),
            "display-style fraction scanner must also require a command boundary");
    }

    @Test
    void explicitFlatParenFenceCanRenderAsVectorText() throws IOException {
        byte[] wmf = VectorWmfFormulaRenderer.render("\\left ( { 第+十+一+届+华+杯+赛 } \\right )", 141.0d, 19.0d);
        List<Integer> records = records(wmf);

        assertTrue(records.contains(0x0A32));
        assertFalse(records.contains(0x0F43));

        byte[] twoRowArray = VectorWmfFormulaRenderer.render("\\begin{array}{c}1\\\\2\\end{array}", 16.0d,
            MathTypeStructureMetrics.ARRAY_SOURCE_HEIGHT_PT);
        assertEquals((int) (MathTypeStructureMetrics.ARRAY_SOURCE_HEIGHT_PT * 20.0d), windowExtY(twoRowArray));
        assertTrue(maxTextYCoordinate(twoRowArray) <= MathTypeStructureMetrics.ARRAY_SOURCE_HEIGHT_PT * 20.0d);
    }

    @Test
    void squarePlaceholderCanRenderAsVectorText() throws IOException {
        byte[] wmf = VectorWmfFormulaRenderer.render("+2=\\square", 40.0d, 13.0d);
        List<Integer> records = records(wmf);

        assertTrue(records.contains(0x0A32));
        assertFalse(records.contains(0x0F43));
    }

    @Test
    void boxedTextCanRenderAsVectorTextAndLines() throws IOException {
        byte[] wmf = VectorWmfFormulaRenderer.render("\\boxed{?????????????}", 92.0d, 18.0d);
        List<Integer> records = records(wmf);

        assertTrue(VectorWmfFormulaRenderer.canRender("\\boxed{?????????????}"));
        assertTrue(VectorWmfFormulaRenderer.canRender("\\boxed{}\\ast (19\\ast 99)=80"));
        assertTrue(VectorWmfFormulaRenderer.canRender("\\overline{bc}\\times a=\\boxed{}5\\boxed{}"));
        assertTrue(VectorWmfFormulaRenderer.canRender("2010\\left| \\overline{abcd}-\\overline{efgh}\\right."));
        assertTrue(VectorWmfFormulaRenderer.canRender("N=\\overline{a3}^{1}\\mathbf{\\times }\\overline{b1}^{8}"));
        assertTrue(VectorWmfFormulaRenderer.canRender("20+\\overline{\\boxed{}\\boxed{}}+8=28+\\overline{\\boxed{}\\boxed{}}"));
        assertTrue(records.contains(0x0325));
        assertTrue(records.contains(0x0A32));
        assertFalse(records.contains(0x0F43));
    }

    @Test
    void simpleArrayCanRenderAsVectorText() throws IOException {
        byte[] wmf = VectorWmfFormulaRenderer.render("\\begin{array}{ccccc} ABCD-EFGH2008 & \\end{array}", 89.0d, 49.0d);
        List<Integer> records = records(wmf);

        assertTrue(VectorWmfFormulaRenderer.canRender("\\begin{array}{c}1\\\\2\\end{array}"));
        assertTrue(VectorWmfFormulaRenderer.canRender("\\left(\\begin{array}{cc} 4A & 8 \\end{array}\\right)"));
        assertTrue(VectorWmfFormulaRenderer.canRender(
            "\\left.\\begin{array}{c} 15| n\\\\ 17| \\left(n+1\\right) \\end{array}"
                + "\\begin{array}{c} \\rightarrow \\\\ \\rightarrow \\end{array}\\right\\}\\Rightarrow [15,17]| \\left(2n-15\\right)"
        ));
        assertTrue(VectorWmfFormulaRenderer.canRender("\\begin{aligned} V&=abh \\\\ V&=Sh \\end{aligned}"));
        assertTrue(VectorWmfFormulaRenderer.canRender(
            "\\begin{array}{l} =\\left(1234+8766\\right)^{2}=10000^{2}\\\\ =100000000 \\end{array}"
        ));
        assertTrue(VectorWmfFormulaRenderer.canRender(
            "\\begin{array}{l} 2x&=6\\\\ x&=3\\\\ \\begin{cases} x=3\\\\ y=2 \\end{cases} \\end{array}"
        ));
        assertTrue(VectorWmfFormulaRenderer.canRender(
            "\\begin{array}{l} \\text{相遇{\\blacksquare}{\\blacksquare}}\\begin{cases} ? 路程=速度和\\times "
                + "\\text{相遇{\\blacksquare}{\\blacksquare}}\\\\ 速度和=? 路程\\div "
                + "\\text{相遇{\\blacksquare}{\\blacksquare}} \\end{cases} \\end{array}"
        ));
        assertTrue(VectorWmfFormulaRenderer.canRender(
            "\\begin{array}{cccccccc} ( & \\boxed{} & + & \\boxed{} & )\\div & \\boxed{} & = & 2\\\\ "
                + "& + & & - & & \\div & & \\\\ & \\boxed{} & - & \\boxed{} & - & \\boxed{} & = & 0\\\\ "
                + "& - & & - & & - & & \\\\ & \\boxed{} & - & \\boxed{} & \\times & \\boxed{} & = & 0\\\\ "
                + "& - & & + & & \\div & & \\\\ & \\boxed{} & + & \\boxed{} & \\div & \\boxed{} & = & 8\\\\ "
                + "& | | & & | | & & | | & & \\\\ & 1 & & 2 & & 6 & & \\end{array}"
        ));
        assertTrue(VectorWmfFormulaRenderer.canRender(
            "\\begin{cases} \\square \\times \\square =5\\square \\\\ \\boxed{12}+\\square =\\square +\\square \\end{cases}"
        ));
        assertTrue(VectorWmfFormulaRenderer.canRender(
            "\\begin{cases} a=4\\left(x-a\\right)+2\\\\ a+8=6\\left(x-a-8\\right) \\end{cases} \\Rightarrow x=147"
        ));
        assertTrue(VectorWmfFormulaRenderer.canRender(
            "\\begin{array}{l} 21\\xrightarrow{红}42\\xrightarrow{红}84\\xrightarrow{红}168\\xrightarrow{红}336"
                + "\\xrightarrow{黄}33\\xrightarrow{黄}3\\\\ "
                + "21\\xrightarrow{红}42\\xrightarrow{红}84\\xrightarrow{红}168\\xrightarrow{黄}16"
                + "\\xrightarrow{红}32\\xrightarrow{黄}3\\\\ "
                + "21\\xrightarrow{红}42\\xrightarrow{黄}4\\xrightarrow{红}8\\xrightarrow{红}16"
                + "\\xrightarrow{红}32\\xrightarrow{黄}3 \\end{array}"
        ));
        assertTrue(VectorWmfFormulaRenderer.canRender(
            "21\\xrightarrow{黄}2\\xrightarrow{红}4\\xrightarrow{红}8\\xrightarrow{红}16"
                + "\\xrightarrow{红}32\\xrightarrow{黄}3"
        ));
        assertTrue(VectorWmfFormulaRenderer.canRender(
            "\\underrightarrow{\\text{A}}\\begin{cases} 6\\\\ 14\\\\ 4 \\end{cases}"
        ));
        assertTrue(VectorWmfFormulaRenderer.canRender(
            "8\\begin{array}{c} \\end{array}8\\begin{array}{c} \\end{array}8=1000"
        ));
        assertTrue(VectorWmfFormulaRenderer.canRender(
            "\\begin{array}{l} 2\\bottom left{12}\\\\ \\begin{array}{l} \\end{array}2\\bottom left{6}\\\\ \\begin{array}{l} 3 \\end{array} \\end{array}"
        ));
        assertTrue(VectorWmfFormulaRenderer.canRender("\\xcancel{\\begin{array}{ccc} & & \\\\ & & \\end{array}}"));
        assertTrue(VectorWmfFormulaRenderer.canRender(
            "\\begin{array}{l} 2\\bottom left{12}\\\\ \\begin{array}{l} \\end{array}2\\bottom left{6}\\\\ "
                + "\\begin{array}{l} \\end{array}\\begin{array}{l} \\end{array}\\begin{array}{l} 3 \\end{array} \\end{array}"
        ));
        assertTrue(VectorWmfFormulaRenderer.canRender(
            "\\begin{array}{l} 2\\bottom left{\\begin{array}{ll} 18 & 12 \\end{array}}\\\\ "
                + "\\begin{array}{l} \\end{array}3\\bottom left{\\begin{array}{ll} 9 & 6 \\end{array}}\\\\ "
                + "\\begin{array}{l} \\end{array}\\begin{array}{l} \\end{array}\\begin{array}{ll} 3 & 2 \\end{array} \\end{array}"
        ));
        assertTrue(records.contains(0x0A32));
        assertFalse(records.contains(0x0F43));
    }

    @Test
    void textCommandsCanRenderAsVectorText() throws IOException {
        byte[] wmf = VectorWmfFormulaRenderer.render("\\vartriangle =\\mathrm{9}+\\cdots", 70.0d, 13.0d);
        List<Integer> records = records(wmf);

        assertTrue(VectorWmfFormulaRenderer.canRender("\\mathrm{9}"));
        assertTrue(VectorWmfFormulaRenderer.canRender("\\boldsymbol{\\pi}"));
        assertTrue(VectorWmfFormulaRenderer.canRender("\\square +\\Circle +\\bigtriangleup +\\mathrm{\\whitestar }="));
        assertTrue(VectorWmfFormulaRenderer.canRender("M\\ast N"));
        assertTrue(VectorWmfFormulaRenderer.canRender("a\\& b=a+b\\div 10"));
        assertTrue(VectorWmfFormulaRenderer.canRender("10.5\\%"));
        assertTrue(VectorWmfFormulaRenderer.canRender("\\{y\\}+y=20.09"));
        assertTrue(VectorWmfFormulaRenderer.canRender("\\left\\{a\\right\\}"));
        assertTrue(VectorWmfFormulaRenderer.canRender("\\left.25\\right| \\left(n+7\\right)"));
        assertTrue(VectorWmfFormulaRenderer.canRender("(3x-2)\\colon (2x+3)=4\\colon 7"));
        assertTrue(VectorWmfFormulaRenderer.canRender("x\\Delta y=\\frac{6\\times x\\times y}{x+2y}"));
        assertTrue(VectorWmfFormulaRenderer.canRender("\\vartriangle"));
        assertTrue(VectorWmfFormulaRenderer.canRender("{\\bigtriangleup}"));
        assertTrue(VectorWmfFormulaRenderer.canRender("{\\bigtriangledown}"));
        assertTrue(VectorWmfFormulaRenderer.canRender("\\uppi =3.14"));
        assertTrue(VectorWmfFormulaRenderer.canRender("1\\sim 9"));
        assertTrue(VectorWmfFormulaRenderer.canRender("{\\backsim}"));
        assertTrue(VectorWmfFormulaRenderer.canRender("\\nmid"));
        assertTrue(VectorWmfFormulaRenderer.canRender("a\\in \\left(10,20\\right)"));
        assertTrue(VectorWmfFormulaRenderer.canRender("x\\notin A\\cup B"));
        assertTrue(VectorWmfFormulaRenderer.canRender("A\\subseteq B\\cap C"));
        assertTrue(VectorWmfFormulaRenderer.canRender("CN\\parallel AF"));
        assertTrue(VectorWmfFormulaRenderer.canRender("BO\\bot AE"));
        assertTrue(VectorWmfFormulaRenderer.canRender("{\\angle}"));
        assertTrue(VectorWmfFormulaRenderer.canRender("{\\llcorner}"));
        assertTrue(VectorWmfFormulaRenderer.canRender("{\\equiv}"));
        assertTrue(VectorWmfFormulaRenderer.canRender("\\left(1\\prec x\\prec \\left(n-1\\right)\\right)"));
        assertTrue(VectorWmfFormulaRenderer.canRender("{\\therefore}"));
        assertTrue(VectorWmfFormulaRenderer.canRender("\\textbf{\\textcolor{maroon}{=20,}}"));
        assertTrue(VectorWmfFormulaRenderer.canRender("0.\\dot{1}+0.125+0.\\dot{3}+0.1\\dot{6}\\approx 0.736"));
        assertTrue(VectorWmfFormulaRenderer.canRender("\\bigcirc"));
        assertTrue(VectorWmfFormulaRenderer.canRender("{\\Circle}"));
        assertTrue(VectorWmfFormulaRenderer.canRender("{\\CIRCLE}"));
        assertTrue(VectorWmfFormulaRenderer.canRender("{\\Sun}"));
        assertTrue(VectorWmfFormulaRenderer.canRender("{\\oplus}"));
        assertTrue(VectorWmfFormulaRenderer.canRender("{\\odot}"));
        assertTrue(VectorWmfFormulaRenderer.canRender("{\\rightarrow}"));
        assertTrue(VectorWmfFormulaRenderer.canRender("\\downarrow"));
        assertTrue(VectorWmfFormulaRenderer.canRender("\\uparrow"));
        assertTrue(VectorWmfFormulaRenderer.canRender("{\\blacksquare}"));
        assertTrue(VectorWmfFormulaRenderer.canRender("{\\Diamond}"));
        assertTrue(VectorWmfFormulaRenderer.canRender("{\\Diamondblack}"));
        assertTrue(VectorWmfFormulaRenderer.canRender("{\\bigstar}"));
        assertTrue(VectorWmfFormulaRenderer.canRender("(1+2+3+\\cdots +9)\\div 3=15"));
        assertTrue(records.contains(0x0A32));
        assertFalse(records.contains(0x0F43));
    }

    @Test
    void simpleScriptsCanRenderAsVectorText() throws IOException {
        byte[] wmf = VectorWmfFormulaRenderer.render("C\\times D=kD ^ { 2 }", 92.0d, 13.0d);
        List<Integer> records = records(wmf);

        assertTrue(VectorWmfFormulaRenderer.canRender("a_{ 1 }"));
        assertTrue(VectorWmfFormulaRenderer.canRender("p_{\\max }"));
        assertTrue(VectorWmfFormulaRenderer.canRender("C\\times D=kD ^ { 2 }"));
        assertTrue(VectorWmfFormulaRenderer.canRender("\\left(n-1\\right)^{2}"));
        assertTrue(VectorWmfFormulaRenderer.canRender("\\left(a\\pm b\\right)^{2}=a^{2}\\pm 2ab+b^{2}"));
        assertTrue(VectorWmfFormulaRenderer.canRender("a\\nabla n=a^{n}+a^{n-1}"));
        assertTrue(VectorWmfFormulaRenderer.canRender("S=(V_{\\mathrm{顺}}+V_{\\mathrm{逆}})\\times 10\\min"));
        assertTrue(VectorWmfFormulaRenderer.canRender("{a_{1}}^{，}=1"));
        assertTrue(VectorWmfFormulaRenderer.canRender("x=2^{m},\\,y=2^{n},"));
        assertTrue(VectorWmfFormulaRenderer.canRender("=\\frac{1}{8}S_{\\euro{}ABCD}=\\frac{1}{8}"));
        assertTrue(VectorWmfFormulaRenderer.canRender("^{\\prime}"));
        assertTrue(records.contains(0x02FB));
        assertTrue(records.contains(0x012D));
        assertTrue(records.contains(0x0A32));
        assertFalse(records.contains(0x0F43));
    }

    @Test
    void simpleFractionsCanRenderAsVectorTextAndLines() throws IOException {
        byte[] wmf = VectorWmfFormulaRenderer.render("\\frac { 1 } { 15 }", 24.0d, 28.0d);
        List<Integer> records = records(wmf);

        assertTrue(VectorWmfFormulaRenderer.canRender("\\frac { 1 } { 15 }"));
        assertTrue(VectorWmfFormulaRenderer.canRender("\\frac{}{}\\:"));
        assertTrue(VectorWmfFormulaRenderer.canRender(
            "\\left ( { 1+2+\\cdots +9+a+b+c } \\right )\\div 3=15+\\frac { a+b+c } { 3 }"
        ));
        assertTrue(VectorWmfFormulaRenderer.canRender(
            "\\frac{1}{9}\\times \\overline{\\underset{1997\\mathrm{个}4}{\\underbrace{444\\cdots 4} }3"
                + "\\underset{1997\\mathrm{个}5}{\\underbrace{555\\cdots 5} }6}"
        ));
        assertTrue(records.contains(0x02FA), "fraction vector WMF should create a pen");
        assertTrue(records.contains(0x0325), "fraction vector WMF should draw a fraction bar");
        assertTrue(records.contains(0x0A32));
        assertFalse(records.contains(0x0F43));
    }

    @Test
    void compactInlineFractionsUseReadableTextHeightWithTightAdvance() throws IOException {
        byte[] inline = VectorWmfFormulaRenderer.render("CO=\\frac{5}{3}", 31.0d, 26.0d);
        byte[] inlineWide = VectorWmfFormulaRenderer.render("48\\times \\frac{1}{4}=12", 50.25d, 27.75d);
        byte[] contextualCfrac = VectorWmfFormulaRenderer.render("x=\\cfrac{a}{b}", 34.0d, 28.0d);
        byte[] standalone = VectorWmfFormulaRenderer.render("\\frac{5}{3}", 16.0d, 26.0d);

        assertTrue(minSelectedFontHeightTwips(inline) >= 200,
            "compact inline fraction slots should not be forced through 8pt script fonts");
        assertTrue(maxTextYCoordinate(inline) - minTextYCoordinate(inline) >= 13.0d * 20.0d,
            "compact inline fraction should use readable numerator/denominator baseline spacing");
        assertEquals(0.965d, VectorWmfFormulaRenderer.compactInlineFractionWidthScale("48\\times \\frac{1}{4}=12"));
        assertEquals(1.0d, VectorWmfFormulaRenderer.compactInlineFractionWidthScale("\\frac{5}{3}"));
        assertEquals(1.0d, VectorWmfFormulaRenderer.compactInlineFractionWidthScale("\\displaystyle\\frac{5}{3}"));
        assertEquals(1.0d, VectorWmfFormulaRenderer.compactInlineFractionWidthScale("\\dfrac{5}{3}"));
        assertEquals(1.0d, VectorWmfFormulaRenderer.compactInlineFractionWidthScale("\\cfrac{5}{3}"));
        assertEquals(1.0d, VectorWmfFormulaRenderer.compactInlineFractionWidthScale("S=\\frac{\\frac{1}{2}}{3}"));
        assertTrue(maxTextYCoordinate(contextualCfrac) > 20.0d * 20.0d,
            "contextual \\cfrac should keep display-style fraction geometry, not compact inline spacing");
        assertTrue(maxTextRightCoordinate(inlineWide) < 49.0d * 20.0d,
            "compact inline fraction rows should no longer fill the whole Word box when MathType reference is narrower");
        assertTrue(textDxTotal(inline, "5") < textDxTotal(standalone, "5"),
            "compact inline fraction should preserve tight digit advance after font-height promotion");
        assertTrue(records(inline).contains(0x0325));
        assertFalse(records(inline).contains(0x0F43));
    }

    @Test
    void nonCompactFractionsUseReadableVerticalSlotSpacing() throws IOException {
        byte[] wmf = VectorWmfFormulaRenderer.render("S=\\frac{\\frac{1}{2}}{3}", 42.0d, 28.0d);
        int baselineSpread = maxTextYCoordinate(wmf) - minTextYCoordinate(wmf);

        assertTrue(baselineSpread >= 280,
            "non-compact fraction numerator and denominator baselines should not be visually cramped: "
                + baselineSpread);
        assertTrue(maxTextYCoordinate(wmf) <= 28.0d * 20.0d);
        assertTrue(maxTextRightCoordinate(wmf) <= 42.0d * 20.0d);
    }

    @Test
    void fractionScriptAreaChainsUseScopedWidthCompression() throws IOException {
        String areaChain = "S_{\\bigtriangleup ENF}=\\frac{9}{9+12+12+16}S_{\\text{梯形EFCD}}"
            + "=\\frac{9}{49}\\times \\frac{7}{12}S=\\frac{3}{28}S";
        String compactAreaChain = "S_{\\bigtriangleup ABC}=\\frac{12}{12+16}S_{\\text{梯形ABCD}}";
        byte[] compressed = VectorWmfFormulaRenderer.render(areaChain, 215.25d, 27.75d);
        byte[] compactArea = VectorWmfFormulaRenderer.render(compactAreaChain, 165.0d, 27.75d);
        byte[] compactInline = VectorWmfFormulaRenderer.render("48\\times \\frac{1}{4}=12", 50.25d, 27.75d);

        assertEquals(0.98d, VectorWmfFormulaRenderer.fractionScriptChainWidthScale(areaChain));
        assertEquals(1.0d, VectorWmfFormulaRenderer.fractionScriptChainWidthScale("48\\times \\frac{1}{4}=12"));
        assertEquals(1.0d, VectorWmfFormulaRenderer.fractionScriptChainWidthScale(
            "EF=\\frac{1}{2}\\left(a+2a\\right)=\\frac{3}{2}a"));
        assertEquals(1.0d, VectorWmfFormulaRenderer.fractionScriptChainWidthScale(
            "S_{n}=\\frac{a_{1}+a_{2}+a_{3}+a_{4}+a_{5}}{5}+b^{2}"));
        assertEquals(0.98d, VectorWmfFormulaRenderer.fractionScriptChainWidthScale(
            compactAreaChain));
        assertEquals(0.965d, VectorWmfFormulaRenderer.compactInlineFractionWidthScale(compactAreaChain));
        assertTrue(maxTextRightCoordinate(compressed) < 212.0d * 20.0d);
        assertTrue(maxTextRightCoordinate(compactArea) > 160.0d * 20.0d,
            "area-chain width compression should not be multiplied by compact-inline compression");
        assertTrue(maxTextRightCoordinate(compactInline) < 49.0d * 20.0d);
    }

    @Test
    void nestedInlineFractionsDoNotUseCompactFractionSlotGeometry() throws IOException {
        byte[] nested = VectorWmfFormulaRenderer.render("CO=\\frac{\\frac{1}{2}}{3}", 42.0d, 28.0d);

        assertTrue(VectorWmfFormulaRenderer.canRender("CO=\\frac{\\frac{1}{2}}{3}"));
        assertTrue(maxTextYCoordinate(nested) > 20.0d * 20.0d,
            "nested inline fractions need non-compact vertical geometry to avoid slot overlap");
        assertTrue(records(nested).contains(0x0325));
        assertFalse(records(nested).contains(0x0F43));
    }

    @Test
    void compactInlineFractionPreservesRealScriptChildren() throws IOException {
        byte[] wmf = VectorWmfFormulaRenderer.render("CO=\\frac{S_{1}}{ABC}", 46.0d, 28.0d);

        assertTrue(textFontHeightTwips(wmf, "S") >= 200,
            "main numerator text should use readable compact fraction height");
        assertTrue(textFontHeightTwips(wmf, "1") < 200,
            "real script children inside compact fraction slots should remain script-sized");
        assertTrue(records(wmf).contains(0x0325));
        assertFalse(records(wmf).contains(0x0F43));
    }

    @Test
    void longLinearFractionSeriesCanRenderAsVectorTextAndLines() throws IOException {
        String latex = "\\left(\\frac{1}{2}+\\frac{1}{3}+\\frac{1}{4}+\\cdots +\\frac{1}{20}\\right)"
            + "+\\left(\\frac{2}{3}+\\frac{2}{4}+\\frac{2}{5}+\\cdots +\\frac{2}{20}\\right)"
            + "+\\frac{19}{20}";
        byte[] wmf = VectorWmfFormulaRenderer.render(latex, 316.0d, 28.0d);
        List<Integer> records = records(wmf);

        assertTrue(VectorWmfFormulaRenderer.canRender(latex));
        assertTrue(VectorWmfFormulaRenderer.canRender(
            "\\begin{array}{l} 5\\mathrm{个篮球}\\,\\,\\,3\\mathrm{个排球}\\,\\,\\,318\\mathrm{元}\\\\ "
                + "\\underline{\\times 2}\\\\ 10\\mathrm{个篮球}\\,\\,6\\mathrm{个排球}\\,\\,\\,636\\mathrm{元} \\end{array}"
        ));
        assertTrue(records.contains(0x0325));
        assertTrue(records.contains(0x0A32));
        assertFalse(records.contains(0x0F43));
    }

    @Test
    void productFractionSeriesWithLdotsCanRenderAsVectorTextAndLines() throws IOException {
        String latex = "2008\\times \\frac{1}{2}\\times \\frac{2}{3}\\times \\frac{3}{4}"
            + "\\times \\ldots \\ldots \\times \\frac{1999}{2000}=\\frac{2008}{2000}=\\frac{251}{250}";
        byte[] wmf = VectorWmfFormulaRenderer.render(latex, 240.0d, 30.0d);
        List<Integer> records = records(wmf);

        assertTrue(VectorWmfFormulaRenderer.canRender(latex));
        assertTrue(records.contains(0x0325));
        assertTrue(records.contains(0x0A32));
        assertFalse(records.contains(0x0F43));
    }

    @Test
    void fractionsWithSimpleScriptsCanRenderAsVectorTextAndLines() throws IOException {
        String latex = "\\frac { a_{ 1 } +a_{ 2 } +a_{ 3 } } { 3 }";
        byte[] wmf = VectorWmfFormulaRenderer.render(latex, 54.0d, 28.0d);
        List<Integer> records = records(wmf);

        assertTrue(VectorWmfFormulaRenderer.canRender(latex));
        assertTrue(VectorWmfFormulaRenderer.canRender("a=\\frac { k ^ { 2 } } { b-k }+k"));
        assertTrue(VectorWmfFormulaRenderer.canRender(
            "=\\frac{1}{4}\\times 20^{2}\\times 21^{2}-8\\times \\frac{1}{4}\\times 10^{2}\\times 11^{2}"
        ));
        assertTrue(VectorWmfFormulaRenderer.canRender(
            "a_{n}=\\frac{\\left(n+1\\right)^{2}}{\\left(n+1+1\\right)\\left(n+1-1\\right)}=\\frac{\\left(n+1\\right)^{2}}{n\\left(n+2\\right)}"
        ));
        assertTrue(VectorWmfFormulaRenderer.canRender("\\frac{\\angle B+\\angle C}{360{^{\\circ}}}\\times \\uppi \\times 2^{2}=4"));
        assertTrue(VectorWmfFormulaRenderer.canRender("\\frac{\\mathrm{P}_{n}^{n}}{n}=\\mathrm{P}_{n-1}^{n-1}"));
        assertTrue(VectorWmfFormulaRenderer.canRender("\\tan \\angle 1=\\frac{DA}{DC}=\\frac{1}{3}"));
        assertTrue(VectorWmfFormulaRenderer.canRender("P_{3}^{1}\\spot P_{5}^{1}"));
        assertTrue(records.contains(0x02FA));
        assertTrue(records.contains(0x0325));
        assertTrue(records.contains(0x0A32));
        assertFalse(records.contains(0x0F43));
    }

    @Test
    void nestedMixedFractionsCanRenderAsVectorTextAndLines() throws IOException {
        String latex = "\\frac{3\\frac{3}{4}\\times 0.2}{1.38}\\times 5.84";
        byte[] wmf = VectorWmfFormulaRenderer.render(latex, 92.0d, 32.0d);
        List<Integer> records = records(wmf);

        assertTrue(VectorWmfFormulaRenderer.canRender(latex));
        assertTrue(VectorWmfFormulaRenderer.canRender(
            "\\frac{\\frac{7}{18}\\times 4.5+0.1\\dot{6}}{13\\frac{1}{3}-3.75\\times 3.2}\\times \\left(\\frac{1}{3}+\\frac{1}{15}+\\frac{1}{35}+\\frac{1}{63}\\right)="
        ));
        assertTrue(VectorWmfFormulaRenderer.canRender(
            "\\frac{\\left(0.\\overset{\\cdot }{6}\\Uptheta \\frac{15}{23}\\right)+\\left(0.625\\Delta \\frac{23}{35}\\right)}{\\left(0.\\overset{\\cdot }{3}\\Delta \\frac{34}{99}\\right)+\\left(\\frac{11}{6}\\Uptheta 2.25\\right)}"
        ));
        assertTrue(VectorWmfFormulaRenderer.canRender(
            "3-\\frac{2}{3-\\frac{2}{\\begin{array}{l} \\enspace \\enspace \\vdots \\\\ 3-\\frac{2}{3} \\end{array}}}"
        ));
        assertTrue(records.contains(0x0325));
        assertTrue(records.contains(0x0A32));
        assertFalse(records.contains(0x0F43));
    }

    @Test
    void annotatedBraceFractionsCanRenderAsVectorTextAndLines() throws IOException {
        String latex = "\\frac{1}{19}+\\frac{202}{1919}+\\frac{30303}{191919}"
            + "+\\frac{\\overset{8\\mathrm{个}90}{\\overbrace{90\\cdots 90} }9}"
            + "{\\underset{9\\mathrm{个}19}{\\underbrace{19\\cdots 19} }}";
        byte[] wmf = VectorWmfFormulaRenderer.render(latex, 188.0d, 34.0d);
        List<Integer> records = records(wmf);

        assertTrue(VectorWmfFormulaRenderer.canRender(latex));
        assertTrue(records.contains(0x0325));
        assertTrue(records.contains(0x0A32));
        assertFalse(records.contains(0x0F43));
    }

    @Test
    void leftBraceArraysCanRenderAsVectorText() throws IOException {
        String latex = "\\left \\{ \\begin{array}{l}B=2 \\\\,s=14\\end{array} \\right.";
        String arrayArrow = "\\begin{array}{c}2H_{2}+O_{2}\\rightarrow 2H_{2}O\\end{array}";
        byte[] wmf = VectorWmfFormulaRenderer.render(latex, 42.0d, 28.0d);
        byte[] arrowWmf = VectorWmfFormulaRenderer.render(arrayArrow, 132.0d, 24.0d);
        List<Integer> records = records(wmf);
        List<Integer> arrowRecords = records(arrowWmf);

        assertTrue(VectorWmfFormulaRenderer.canRender(latex));
        assertTrue(VectorWmfFormulaRenderer.canRender(arrayArrow));
        assertTrue(VectorWmfFormulaRenderer.canRender("\\begin{cases} b=1\\\\ c=1 \\end{cases}"));
        assertTrue(VectorWmfFormulaRenderer.canRender(
            "\\begin{cases} x=1\\\\ y=3 \\end{cases} ,\\begin{cases} x=6\\\\ y=1 \\end{cases}"
        ));
        assertTrue(VectorWmfFormulaRenderer.canRender(
            "\\begin{array}{l} \\begin{cases} \\frac{x}{y}=\\frac{7}{3}\\\\ "
                + "\\frac{x+70}{y+70}=\\frac{7}{4} \\end{cases} \\\\ \\end{array}"
        ));
        assertTrue(records.contains(0x02FB));
        assertTrue(records.contains(0x0A32));
        assertTrue(arrowRecords.contains(0x0325));
        assertFalse(records.contains(0x0F43));
        assertFalse(arrowRecords.contains(0x0F43));
    }

    @Test
    void emptyAndParenArraysCanRenderAsVectorText() throws IOException {
        String empty = "\\begin{array}{cccc} {} & \\end{array}";
        String paren = "\\left ( { \\begin{array}{cc} {} & 4, \\\\,8 & \\end{array} } \\right )";

        assertTrue(VectorWmfFormulaRenderer.canRender(empty));
        assertTrue(VectorWmfFormulaRenderer.canRender(paren));
        assertFalse(records(VectorWmfFormulaRenderer.render(empty, 18.0d, 13.0d)).contains(0x0F43));
        assertFalse(records(VectorWmfFormulaRenderer.render(paren, 34.0d, 33.0d)).contains(0x0F43));
    }

    @Test
    void adjacentArraysCanRenderAsVectorText() throws IOException {
        String adjacent = "\\begin{array}{l}cba \\\\,\\times abc\\end{array}\\begin{array}{l}c \\\\,b \\\\,b \\\\,a\\end{array}";
        String withText = "\\begin{array}{l}1b5 \\\\,\\times 5b1\\end{array}1b505\\begin{array}{l}1 \\\\,b \\\\,b \\\\,5\\end{array}";

        assertTrue(VectorWmfFormulaRenderer.canRender(adjacent));
        assertTrue(VectorWmfFormulaRenderer.canRender(withText));
        assertFalse(records(VectorWmfFormulaRenderer.render(adjacent, 60.0d, 50.0d)).contains(0x0F43));
        assertFalse(records(VectorWmfFormulaRenderer.render(withText, 92.0d, 50.0d)).contains(0x0F43));
    }

    @Test
    void widePuzzleArraysCanRenderAsVectorText() throws IOException {
        String latex = "\\begin{array}{cccccccc} ( & \\mathrm{7} & + & \\mathrm{9} & )\\div & \\mathrm{8} & = & \\mathrm{2} \\\\,{} & + & {} & - & {} & \\div & {} & {} \\\\,{} & \\mathrm{1}\\mathrm{1} & - & \\mathrm{1}\\mathrm{0} & - & \\mathrm{1} & = & \\mathrm{0} \\\\,{} & - & {} & - & {} & - & {} & {} \\\\,{} & \\mathrm{1}\\mathrm{2} & - & \\mathrm{3} & \\times & \\mathrm{4} & = & \\mathrm{0} \\\\,{} & - & {} & + & {} & \\div & {} & {} \\\\,{} & \\mathrm{5} & + & \\mathrm{6} & \\div & \\mathrm{2} & = & \\mathrm{8} \\\\,{} & \\|\\| & {} & \\|\\| & {} & \\|\\| & {} & {} \\\\,{} & \\mathrm{1} & {} & \\mathrm{2} & {} & \\mathrm{6} & {} & {} \\end{array}";

        assertTrue(VectorWmfFormulaRenderer.canRender(latex));
        assertFalse(records(VectorWmfFormulaRenderer.render(latex, 150.0d, 110.0d)).contains(0x0F43));
    }

    @Test
    void hlineArraysCanRenderAsVectorTextAndLines() throws IOException {
        String latex = "\\begin{array}{cccccccccccccccccc} & \\mathrm{和} & = & 1 & + & 2 & + & 3 & + & 4 & + & \\cdots & + & 98 & + & 99 & + & 100\\\\ + & \\mathrm{和} & = & 100 & + & 99 & + & 98 & + & 97 & + & \\cdots & + & 3 & + & 2 & + & 1\\\\ \\hline & 2\\mathrm{倍和} & = & 101 & + & 101 & + & 101 & + & 101 & + & \\cdots & + & 101 & + & 101 & + & 101 \\end{array}";
        List<Integer> records = records(VectorWmfFormulaRenderer.render(latex, 260.0d, 42.0d));

        assertTrue(VectorWmfFormulaRenderer.canRender(latex));
        assertTrue(records.contains(0x0325));
        assertTrue(records.contains(0x0A32));
        assertFalse(records.contains(0x0F43));
    }

    @Test
    void overarcCanRenderAsVectorTextAndArcLines() throws IOException {
        String latex = "\\overarc{ACD}=\\overarc{AC}\\times 3=240";
        List<Integer> records = records(VectorWmfFormulaRenderer.render(latex, 120.0d, 18.0d));

        assertTrue(VectorWmfFormulaRenderer.canRender(latex));
        assertTrue(VectorWmfFormulaRenderer.canRender("\\arc{AB}=60"));
        assertTrue(VectorWmfFormulaRenderer.canRender("\\wideparen{ABC}=120"));
        assertTrue(VectorWmfFormulaRenderer.canRender("\\overset{\\frown }{\\mathrm{AEB}}"));
        assertTrue(records.contains(0x0325));
        assertTrue(records.contains(0x0A32));
        assertFalse(records.contains(0x0F43));
    }

    @Test
    void accentLayoutsUseSourceHeightFamily() throws IOException {
        byte[] overline = VectorWmfFormulaRenderer.render("\\overline{AB}", 24.0d,
            MathTypeStructureMetrics.ACCENT_HEIGHT_PT);
        byte[] underline = VectorWmfFormulaRenderer.render("\\underline{AB}", 24.0d,
            MathTypeStructureMetrics.ACCENT_HEIGHT_PT);

        assertTrue(VectorWmfFormulaRenderer.canRender("\\overline{AB}"));
        assertTrue(VectorWmfFormulaRenderer.canRender("\\underline{AB}"));
        assertEquals((int) (MathTypeStructureMetrics.ACCENT_HEIGHT_PT * 20.0d), windowExtY(overline));
        assertEquals((int) (MathTypeStructureMetrics.ACCENT_HEIGHT_PT * 20.0d), windowExtY(underline));
        assertTrue(records(overline).contains(0x0325));
        assertTrue(records(underline).contains(0x0325));
        assertFalse(records(overline).contains(0x0F43));
        assertFalse(records(underline).contains(0x0F43));
        assertTrue(maxPolylineYCoordinate(overline) <= MathTypeStructureMetrics.ACCENT_HEIGHT_PT * 20.0d);
        assertTrue(maxPolylineYCoordinate(underline) <= MathTypeStructureMetrics.ACCENT_HEIGHT_PT * 20.0d);
        assertTrue(maxTextYCoordinate(overline) <= MathTypeStructureMetrics.ACCENT_HEIGHT_PT * 20.0d);
        assertTrue(maxTextYCoordinate(underline) <= MathTypeStructureMetrics.ACCENT_HEIGHT_PT * 20.0d);
    }

    @Test
    void nestedArrayCellsCanRenderAsVectorText() throws IOException {
        String latex = "\\begin{array}{ccccc} {} & {} & 11 & {} & {} \\\\,{} & 17 & {} & 13 & {} \\\\,23 & {} & 19 & {} & 15 \\\\,{} & 25 & {} & 21 & {} \\\\,{} & {} & 27 & {} & {} \\\\,\\to & \\begin{array}{ccccc} {} & \\end{array} & {} & {} & 27 \\\\,{} & {} & {} & 17 & {} \\\\,13 & {} & 23 & {} & 19 \\\\,{} & 15 & {} & 25 & {} \\\\,21 & {} & {} & {} & 11 \\\\,{} & {} & \\to & \\begin{array}{ccccc} {} & \\end{array} & {} \\\\,{} & 27 & {} & {} & {} \\\\,17 & {} & 13 & {} & 15 \\\\,{} & 19 & {} & 23 & {} \\\\,25 & {} & 21 & {} & {} \\\\,{} & 11 & {} & {} & \\to \\\\,\\begin{array}{ccc} 172713151923251121 & \\end{array} & \\end{array}";

        assertTrue(VectorWmfFormulaRenderer.canRender(latex));
        assertFalse(records(VectorWmfFormulaRenderer.render(latex, 168.0d, 205.0d)).contains(0x0F43));
    }

    @Test
    void standaloneScriptsAndEscapedUnderscoresCanRenderAsVectorText() throws IOException {
        byte[] script = VectorWmfFormulaRenderer.render("^ { \\mathrm{3} }", 12.0d, 13.0d);
        byte[] underline = VectorWmfFormulaRenderer.render("EF=\\_\\_\\_\\_\\_", 56.0d, 13.0d);

        assertTrue(VectorWmfFormulaRenderer.canRender("^ { \\mathrm{3} }"));
        assertTrue(VectorWmfFormulaRenderer.canRender("EF=\\_\\_\\_\\_\\_"));
        assertTrue(records(script).contains(0x0A32));
        assertTrue(records(underline).contains(0x0A32));
        assertFalse(records(script).contains(0x0F43));
        assertFalse(records(underline).contains(0x0F43));
    }

    @Test
    void pairedScriptsStayInsidePreviewWidth() throws IOException {
        byte[] wmf = VectorWmfFormulaRenderer.render("x_i^j", 18.0d, 13.0d);
        byte[] superscript = VectorWmfFormulaRenderer.render("a^2", 13.0d, 13.0d);
        byte[] subscript = VectorWmfFormulaRenderer.render("x_i", 13.0d, 13.0d);

        assertTrue(VectorWmfFormulaRenderer.canRender("x_i^j"));
        assertTrue(VectorWmfFormulaRenderer.canRender("a^2"));
        assertTrue(VectorWmfFormulaRenderer.canRender("x_i"));
        assertFalse(records(wmf).contains(0x0F43));
        assertFalse(records(superscript).contains(0x0F43));
        assertFalse(records(subscript).contains(0x0F43));
        assertTrue(maxRecordCoordinate(wmf) <= 18.0d * 20.0d);
        assertTrue(maxRecordCoordinate(superscript) <= 13.0d * 20.0d);
        assertTrue(maxRecordCoordinate(subscript) <= 13.0d * 20.0d);
    }

    @Test
    void onlySimpleShortScriptsUseReducedScriptFontHeight() throws IOException {
        byte[] shortScript = VectorWmfFormulaRenderer.render("a^2", 13.0d, 13.0d);
        byte[] widerScript = VectorWmfFormulaRenderer.render("\\left(a+b\\right)^2", 36.0d, 18.75d);

        assertTrue(createFontHeights(shortScript).get(1) > -160);
        assertEquals(-163, createFontHeights(widerScript).get(1));
        assertTrue(maxRecordCoordinate(shortScript) <= 13.0d * 20.0d);
        assertTrue(maxRecordCoordinate(widerScript) <= 36.0d * 20.0d);
    }

    @Test
    void simpleShortScriptsUseCompactHorizontalAdvance() throws IOException {
        byte[] shortScript = VectorWmfFormulaRenderer.render("a^2", 13.0d, 13.0d);
        byte[] widerScript = VectorWmfFormulaRenderer.render("S=\\left(a+b\\right)^2=9", 76.0d, 18.75d);

        assertTrue(createFontWidths(shortScript).get(1) > 0);
        assertEquals(0, createFontWidths(widerScript).get(1));
        assertTrue(totalTextDx(shortScript) < totalTextDx(widerScript));
        assertTrue(firstTextDxTotal(shortScript) < firstTextDxTotal(widerScript));
        assertTrue(maxRecordCoordinate(shortScript) <= 13.0d * 20.0d);
        assertTrue(maxRecordCoordinate(widerScript) <= 76.0d * 20.0d);
    }

    @Test
    void standaloneShortScriptsCompactWholeRunWithoutChangingRatioFormulas() throws IOException {
        byte[] upperSubscript = VectorWmfFormulaRenderer.render("S_{1}", 12.0d, 15.75d);
        byte[] lowerSuperscript = VectorWmfFormulaRenderer.render("a^{2}", 12.75d, 15.0d);
        byte[] ratioFormula = VectorWmfFormulaRenderer.render("S_{1}\\colon S_{3}=a^{2}\\colon b^{2}", 64.0d, 17.25d);
        byte[] rawColonFormula = VectorWmfFormulaRenderer.render("S_{1}:S_{3}=a^{2}:b^{2}", 64.0d, 17.25d);
        byte[] longChainFormula = VectorWmfFormulaRenderer.render(
            "S_{1}\\colon S_{3}\\colon S_{2}\\colon S_{4}=a^{2}\\colon b^{2}\\colon ab\\colon ab",
            131.25d, 17.25d);
        byte[] triangleChainFormula = VectorWmfFormulaRenderer.render(
            "S_{\\bigtriangleup AOB}\\colon S_{\\bigtriangleup BOC}=a^{2}\\colon ab=25\\colon 35",
            131.25d, 17.25d);
        String longTriangleRelation = "S_{\\bigtriangleup GEF}\\colon S_{\\bigtriangleup GBD}"
            + "\\colon S_{\\bigtriangleup DGF}\\colon S_{nBGE}=EF^{2}\\colon BD^{2}"
            + "\\colon DF^{2}\\colon DE^{2}=8\\colon18\\colon12\\colon12=4\\colon9\\colon6\\colon6";
        String incidentalTriangleLongRelation = "S_{1}\\colon S_{3}\\colon S_{2}\\colon S_{4}=a^{2}"
            + "\\colon b^{2}\\colon ab\\colon ab\\colon \\bigtriangleup X"
            + "\\colon c^{2}\\colon d^{2}\\colon e^{2}\\colon f^{2}";
        byte[] longTriangleChainFormula = VectorWmfFormulaRenderer.render(longTriangleRelation, 368.25d, 17.25d);
        byte[] leftRightEquation = VectorWmfFormulaRenderer.render("S=\\left(a+b\\right)^2=9", 76.0d, 18.75d);
        byte[] repeatedLeftRightEquation = VectorWmfFormulaRenderer.render(
            "S=\\left(a+b\\right)^2=\\left(1+2\\right)^2=9", 111.0d, 18.75d);

        assertTrue(VectorWmfFormulaRenderer.canRender("S_{1}"));
        assertTrue(VectorWmfFormulaRenderer.canRender("a^{2}"));
        assertTrue(VectorWmfFormulaRenderer.isStandaloneUpperSubscript("S_{1}"));
        assertFalse(VectorWmfFormulaRenderer.isStandaloneUpperSubscript("a^{2}"));
        assertFalse(VectorWmfFormulaRenderer.isStandaloneUpperSubscript("S_{1}\\colon S_{3}=a^{2}\\colon b^{2}"));
        assertTrue(VectorWmfFormulaRenderer.canRender("S_{1}\\colon S_{3}=a^{2}\\colon b^{2}"));
        assertTrue(VectorWmfFormulaRenderer.canRender("S_{1}:S_{3}=a^{2}:b^{2}"));
        assertTrue(VectorWmfFormulaRenderer.canRender(
            "S_{1}\\colon S_{3}\\colon S_{2}\\colon S_{4}=a^{2}\\colon b^{2}\\colon ab\\colon ab"));
        assertTrue(VectorWmfFormulaRenderer.canRender(
            "S_{\\bigtriangleup AOB}\\colon S_{\\bigtriangleup BOC}=a^{2}\\colon ab=25\\colon 35"));
        assertTrue(VectorWmfFormulaRenderer.canRender("S=\\left(a+b\\right)^2=9"));
        assertTrue(VectorWmfFormulaRenderer.canRender("a^{-1}b"));
        assertTrue(VectorWmfFormulaRenderer.canRender("a^{-1}+b"));
        assertTrue(maxTextRightCoordinate(upperSubscript) < 12.0d * 20.0d);
        assertTrue(maxTextRightCoordinate(lowerSuperscript) < 12.75d * 20.0d);
        assertEquals(maxTextRightCoordinate(rawColonFormula), maxTextRightCoordinate(ratioFormula));
        assertEquals(totalTextDx(rawColonFormula), totalTextDx(ratioFormula));
        assertFalse(VectorWmfFormulaRenderer.hasTopLevelRelationOperator("a^{-1}b"));
        assertFalse(VectorWmfFormulaRenderer.hasTopLevelRelationOperator("a^{2}(b-c)"));
        assertTrue(VectorWmfFormulaRenderer.hasTopLevelRelationOperator("a^{-1}+b"));
        assertTrue(VectorWmfFormulaRenderer.hasTopLevelRelationOperator("S_{1}:S_{3}=a^{2}:b^{2}"));
        assertEquals(1.0d, VectorWmfFormulaRenderer.scriptRelationWidthScale("a^{-1}b"));
        assertEquals(1.0d, VectorWmfFormulaRenderer.scriptRelationWidthScale("a^{2}(b-c)"));
        assertEquals(0.975d, VectorWmfFormulaRenderer.scriptRelationWidthScale("a^{2}+b"));
        assertEquals(3, VectorWmfFormulaRenderer.topLevelRelationOperatorCount("S_{1}:S_{3}=a^{2}:b^{2}"));
        assertEquals(7, VectorWmfFormulaRenderer.topLevelRelationOperatorCount(
            "S_{1}:S_{3}:S_{2}:S_{4}=a^{2}:b^{2}:ab:ab"));
        assertEquals(0.970d, VectorWmfFormulaRenderer.scriptRelationWidthScale(
            "S_{1}\\colon S_{3}=a^{2}\\colon b^{2}"));
        assertEquals(0.970d, VectorWmfFormulaRenderer.scriptRelationWidthScale(
            "S_{1}:S_{3}=a^{2}:b^{2}"));
        assertEquals(0.9895d, VectorWmfFormulaRenderer.scriptRelationWidthScale(
            "S_{1}\\colon S_{3}\\colon S_{2}\\colon S_{4}=a^{2}\\colon b^{2}\\colon ab\\colon ab"));
        assertEquals(0.98d, VectorWmfFormulaRenderer.scriptRelationWidthScale(
            "S_{\\bigtriangleup AOB}\\colon S_{\\bigtriangleup BOC}=a^{2}\\colon ab=25\\colon 35"));
        assertEquals(0.98d, VectorWmfFormulaRenderer.scriptRelationWidthScale(
            "S_{\\bigtriangleup AOB}\\colon S_{\\bigtriangleup BOC}\\colon S_{\\bigtriangleup COD}"
                + "\\colon S_{\\bigtriangleup DOA}=a^{2}\\colon ab\\colon b^{2}\\colon ac"));
        assertEquals(0.981d, VectorWmfFormulaRenderer.scriptRelationWidthScale(longTriangleRelation));
        assertEquals(0.9895d, VectorWmfFormulaRenderer.scriptRelationWidthScale(incidentalTriangleLongRelation));
        assertEquals(0.890d, VectorWmfFormulaRenderer.scriptRelationFontYScale(
            "S_{1}\\colon S_{3}=a^{2}\\colon b^{2}"));
        assertEquals(0.890d, VectorWmfFormulaRenderer.scriptRelationFontYScale(
            "S_{1}\\colon S_{3}\\colon S_{2}\\colon S_{4}=a^{2}\\colon b^{2}\\colon ab\\colon ab"));
        assertEquals(0.887d, VectorWmfFormulaRenderer.scriptRelationFontYScale(
            "S_{\\bigtriangleup AOB}\\colon S_{\\bigtriangleup BOC}=a^{2}\\colon ab=25\\colon 35"));
        assertEquals(0.955d, VectorWmfFormulaRenderer.scriptRelationFontYScale(longTriangleRelation));
        assertEquals(0.890d, VectorWmfFormulaRenderer.scriptRelationFontYScale(incidentalTriangleLongRelation));
        assertEquals(1.0d, VectorWmfFormulaRenderer.scriptRelationFontYScale("S_{2}=2"));
        assertEquals(0.975d, VectorWmfFormulaRenderer.scriptRelationWidthScale("S_{1}=a^{2}=1"));
        assertEquals(0.975d, VectorWmfFormulaRenderer.scriptRelationWidthScale("S_{3}=4=b^{2}"));
        assertEquals(0.887d, VectorWmfFormulaRenderer.scriptRelationFontYScale("S_{1}=a^{2}=1"));
        assertEquals(0.890d, VectorWmfFormulaRenderer.scriptRelationFontYScale("S_{3}=4=b^{2}"));
        assertEquals(0.93d, VectorWmfFormulaRenderer.scriptRelationFontYScale("S_{2}=2=a\\times b"));
        assertEquals(0.93d, VectorWmfFormulaRenderer.scriptRelationFontYScale(
            "S=S_{1}+S_{2}+S_{3}+S_{4}=1+2+4+2=9"));
        assertEquals(1.0d, VectorWmfFormulaRenderer.scriptRelationFontYScale("a^{-1}b"));
        assertEquals(1.0d, VectorWmfFormulaRenderer.scriptRelationFontYScale("a^{2}+b"));
        assertEquals(1.0d, VectorWmfFormulaRenderer.scriptRelationFontYScale("a^{2}-b"));
        assertEquals(1.113d, VectorWmfFormulaRenderer.scriptRelationHeightWidthCompensation("S_{1}=a^{2}=1"));
        assertEquals(1.025d, VectorWmfFormulaRenderer.scriptRelationHeightWidthCompensation(
            "S_{\\bigtriangleup AOB}\\colon S_{\\bigtriangleup BOC}=a^{2}\\colon ab=25\\colon 35"));
        assertEquals(1.016d, VectorWmfFormulaRenderer.scriptRelationHeightWidthCompensation(longTriangleRelation));
        assertEquals(1.022d,
            VectorWmfFormulaRenderer.scriptRelationHeightWidthCompensation(incidentalTriangleLongRelation));
        assertEquals(1.022d, VectorWmfFormulaRenderer.scriptRelationHeightWidthCompensation(
            "S_{1}\\colon S_{3}\\colon S_{2}\\colon S_{4}=a^{2}\\colon b^{2}\\colon ab\\colon ab"));
        assertEquals(1.0d, VectorWmfFormulaRenderer.scriptRelationHeightWidthCompensation(
            "S=S_{1}+S_{2}+S_{3}+S_{4}=1+2+4+2=9"));
        assertEquals(1.0d, VectorWmfFormulaRenderer.scriptRelationHeightWidthCompensation(
            "S_{1}\\colon S_{3}=a^{2}\\colon b^{2}"));
        assertEquals(1.0d, VectorWmfFormulaRenderer.scriptRelationHeightWidthCompensation("S_{3}=4=b^{2}"));
        assertEquals(1.0d, VectorWmfFormulaRenderer.scriptRelationHeightWidthCompensation("S_{2}=2=a\\times b"));
        assertFalse(VectorWmfFormulaRenderer.repeatedEquationParenPower("S=\\left(a+b\\right)^2=9"));
        assertTrue(VectorWmfFormulaRenderer.repeatedEquationParenPower(
            "S=\\left(a+b\\right)^2=\\left(1+2\\right)^2=9"));
        assertTrue(maxTextRightCoordinate(ratioFormula) < 62.5d * 20.0d);
        assertTrue(maxTextRightCoordinate(longChainFormula) > 126.0d * 20.0d);
        assertTrue(maxTextRightCoordinate(triangleChainFormula) < 129.0d * 20.0d);
        assertTrue(maxTextRightCoordinate(longTriangleChainFormula) > 360.0d * 20.0d);
        assertTrue(maxTextRightCoordinate(longTriangleChainFormula) <= 368.25d * 20.0d);
        assertTrue(maxTextRightCoordinate(ratioFormula) > maxTextRightCoordinate(upperSubscript) * 5);
        assertTrue(totalTextDx(ratioFormula) > totalTextDx(upperSubscript) * 6);
        assertTrue(maxTextRightCoordinate(leftRightEquation) > maxTextRightCoordinate(upperSubscript) * 6);
        assertTrue(maxTextRightCoordinate(repeatedLeftRightEquation) < 2200);
        assertTrue(totalTextDx(repeatedLeftRightEquation) < totalTextDx(leftRightEquation) * 1.455d);
        assertTrue(totalTextDx(repeatedLeftRightEquation) < 2200);
        assertTrue(maxRecordCoordinate(upperSubscript) <= 12.0d * 20.0d);
        assertTrue(maxRecordCoordinate(lowerSuperscript) <= 12.75d * 20.0d);
        assertTrue(maxRecordCoordinate(ratioFormula) <= 64.0d * 20.0d);
        assertTrue(maxRecordCoordinate(rawColonFormula) <= 64.0d * 20.0d);
        assertTrue(maxRecordCoordinate(longChainFormula) <= 131.25d * 20.0d);
        assertTrue(maxRecordCoordinate(longTriangleChainFormula) <= 368.25d * 20.0d);
        assertTrue(maxRecordCoordinate(triangleChainFormula) <= 131.25d * 20.0d);
        assertTrue(maxRecordCoordinate(leftRightEquation) <= 76.0d * 20.0d);
    }

    @Test
    void shortGeometryLabelsUseCompactHorizontalAdvance() throws IOException {
        byte[] geometry = VectorWmfFormulaRenderer.render("ABCD", 36.0d, 12.75d);
        byte[] mixed = VectorWmfFormulaRenderer.render("ABCD1", 36.0d, 12.75d);

        int geometryDx = totalTextDx(geometry);
        int mixedDx = totalTextDx(mixed);

        assertTrue(VectorWmfFormulaRenderer.canRender("ABCD"));
        assertTrue(VectorWmfFormulaRenderer.canRender("ABCD1"));
        assertTrue(geometryDx > 0);
        assertTrue(mixedDx > 0);
        assertTrue(geometryDx < mixedDx);
        assertTrue(maxRecordCoordinate(geometry) <= 36.0d * 20.0d);
    }

    @Test
    void standaloneTwoDigitObjectsUseCompactHorizontalAdvance() throws IOException {
        byte[] twoDigits = VectorWmfFormulaRenderer.render("25", 14.25d, 12.75d);
        byte[] anotherTwoDigits = VectorWmfFormulaRenderer.render("35", 14.25d, 12.75d);
        byte[] threeDigits = VectorWmfFormulaRenderer.render("250", 22.0d, 12.75d);
        byte[] embeddedDigits = VectorWmfFormulaRenderer.render("S=25+35", 64.0d, 12.75d);

        assertTrue(VectorWmfFormulaRenderer.canRender("25"));
        assertTrue(VectorWmfFormulaRenderer.canRender("35"));
        assertTrue(VectorWmfFormulaRenderer.canRender("250"));
        assertTrue(VectorWmfFormulaRenderer.canRender("S=25+35"));
        assertEquals(firstTextAverageDx(twoDigits), firstTextAverageDx(anotherTwoDigits));
        assertTrue(firstTextAverageDx(twoDigits) < firstTextAverageDx(threeDigits));
        assertTrue(firstTextAverageDx(embeddedDigits) > firstTextAverageDx(twoDigits));
        assertTrue(maxRecordCoordinate(twoDigits) <= 14.25d * 20.0d);
        assertTrue(maxRecordCoordinate(anotherTwoDigits) <= 14.25d * 20.0d);
        assertTrue(maxRecordCoordinate(threeDigits) <= 22.0d * 20.0d);
        assertTrue(maxRecordCoordinate(embeddedDigits) <= 64.0d * 20.0d);
    }

    @Test
    void standaloneParenthesizedPowersUseCompactHorizontalAdvance() throws IOException {
        byte[] standalone = VectorWmfFormulaRenderer.render("\\left(a+b\\right)^2", 36.0d, 18.75d);
        byte[] metricsStandalone = VectorWmfFormulaRenderer.render(
            "\\pwmetrics{36.000,19.000,36.000,18.750}\\left(a+b\\right)^{2}", 36.0d, 18.75d);
        byte[] plainParen = VectorWmfFormulaRenderer.render("(a+b)^2", 36.0d, 18.75d);
        byte[] longEquation = VectorWmfFormulaRenderer.render("S=\\left(a+b\\right)^2=9", 76.0d, 18.75d);
        byte[] repeatedEquation = VectorWmfFormulaRenderer.render(
            "S=\\left(a+b\\right)^2=\\left(1+2\\right)^2=9", 111.0d, 18.75d);

        assertTrue(VectorWmfFormulaRenderer.canRender("\\left(a+b\\right)^2"));
        assertTrue(VectorWmfFormulaRenderer.canRender("(a+b)^2"));
        assertTrue(VectorWmfFormulaRenderer.canRender("S=\\left(a+b\\right)^2=9"));
        assertEquals(0, createFontWidths(standalone).get(0));
        assertEquals(0, createFontWidths(standalone).get(1));
        assertEquals(0, createFontWidths(metricsStandalone).get(0));
        assertEquals(0, createFontWidths(metricsStandalone).get(1));
        assertEquals(0, createFontWidths(plainParen).get(0));
        assertEquals(0, createFontWidths(plainParen).get(1));
        assertEquals(0, createFontWidths(longEquation).get(0));
        assertEquals(0, createFontWidths(longEquation).get(1));
        assertEquals(totalTextDx(standalone), totalTextDx(metricsStandalone));
        assertTrue(maxTextRightCoordinate(standalone) < maxTextRightCoordinate(plainParen));
        assertEquals(maxTextRightCoordinate(standalone), maxTextRightCoordinate(metricsStandalone));
        assertFalse(VectorWmfFormulaRenderer.repeatedEquationParenPower("S=\\left(a+b\\right)^2=9"));
        assertTrue(VectorWmfFormulaRenderer.repeatedEquationParenPower(
            "S=\\left(a+b\\right)^2=\\left(1+2\\right)^2=9"));
        assertTrue(VectorWmfFormulaRenderer.hasClosingFenceSuperscript("\\left(a+b\\right)^2"));
        assertTrue(VectorWmfFormulaRenderer.hasClosingFenceSuperscript("S=\\left(a+b\\right)^2=9"));
        assertTrue(VectorWmfFormulaRenderer.hasClosingFenceSuperscript("(a+b)^2"));
        assertFalse(VectorWmfFormulaRenderer.hasClosingFenceSuperscript("S_{1}\\colon S_{3}=a^{2}\\colon b^{2}"));
        assertFalse(VectorWmfFormulaRenderer.hasClosingFenceSuperscript("a^{2}+b"));
        assertTrue(maxTextRightCoordinate(repeatedEquation) < 2189);
        assertTrue(totalTextDx(repeatedEquation) < 2200);
        assertTrue(totalTextDx(standalone) < totalTextDx(longEquation));
        assertTrue(maxRecordCoordinate(standalone) <= 36.0d * 20.0d);
        assertTrue(maxRecordCoordinate(metricsStandalone) <= 36.0d * 20.0d);
        assertTrue(maxRecordCoordinate(plainParen) <= 36.0d * 20.0d);
        assertTrue(maxRecordCoordinate(longEquation) <= 76.0d * 20.0d);
        assertTrue(maxRecordCoordinate(repeatedEquation) <= 111.0d * 20.0d);
    }

    @Test
    void standaloneShortLabelsUseLocalWidthCorrection() throws IOException {
        byte[] singleS = VectorWmfFormulaRenderer.render("\\pwmetrics{9.983,12.977,9.750,12.750}S", 9.75d,
            12.75d);
        byte[] singleA = VectorWmfFormulaRenderer.render("A", 9.75d, 12.75d);
        byte[] scriptS = VectorWmfFormulaRenderer.render("S_{1}", 12.0d, 15.75d);
        byte[] scriptS3 = VectorWmfFormulaRenderer.render("S_{3}", 12.0d, 15.75d);
        byte[] scriptA1 = VectorWmfFormulaRenderer.render("A_{1}", 12.0d, 15.75d);
        byte[] bd = VectorWmfFormulaRenderer.render("\\pwmetrics{17.992,11.995,18.000,12.000}BD", 18.0d,
            12.0d);
        byte[] ab = VectorWmfFormulaRenderer.render("\\pwmetrics{17.992,11.995,18.000,12.000}AB", 18.0d,
            12.0d);
        byte[] abcd = VectorWmfFormulaRenderer.render("ABCD", 32.25d, 12.75d);
        byte[] twentyFive = VectorWmfFormulaRenderer.render("25", 14.25d, 12.75d);
        byte[] thirtyFive = VectorWmfFormulaRenderer.render("35", 14.25d, 12.75d);
        byte[] longNumber = VectorWmfFormulaRenderer.render("144", 20.0d, 12.75d);
        byte[] aSquared = VectorWmfFormulaRenderer.render("a^{2}", 9.75d, 15.75d);
        byte[] bSquared = VectorWmfFormulaRenderer.render("b^{2}", 9.75d, 15.75d);
        byte[] bCubed = VectorWmfFormulaRenderer.render("b^{3}", 9.75d, 15.75d);
        byte[] cSquared = VectorWmfFormulaRenderer.render("c^{2}", 9.75d, 15.75d);

        assertTrue(createFontWidths(singleS).get(0) > 0);
        assertEquals(100, createFontWidths(singleS).get(0));
        assertEquals(0, createFontWidths(singleA).get(0));
        assertEquals(0, createFontWidths(scriptS).get(0));
        assertEquals(0, createFontWidths(scriptS3).get(0));
        assertTrue(firstTextDxTotal(singleS) > firstTextDxTotal(singleA));
        assertTrue(maxTextRightCoordinate(singleS) < 10.5d * 20.0d);
        assertTrue(totalTextDx(scriptS) < totalTextDx(scriptA1));
        assertTrue(totalTextDx(scriptS3) > totalTextDx(scriptS));
        assertTrue(firstTextDxTotal(bd) > firstTextDxTotal(ab));
        assertTrue(firstTextDxTotal(twentyFive) < firstTextDxTotal(longNumber));
        assertEquals(firstTextDxTotal(twentyFive), firstTextDxTotal(thirtyFive));
        assertTrue(maxTextRightCoordinate(bSquared) < maxTextRightCoordinate(aSquared));
        assertTrue(maxTextRightCoordinate(bSquared) < 172);
        assertTrue(maxTextRightCoordinate(bCubed) > maxTextRightCoordinate(bSquared));
        assertTrue(maxTextRightCoordinate(cSquared) > maxTextRightCoordinate(bSquared));
        assertTrue(maxRecordCoordinate(bd) <= 18.0d * 20.0d);
        assertTrue(maxRecordCoordinate(abcd) <= 32.25d * 20.0d);
        assertTrue(maxRecordCoordinate(twentyFive) <= 14.25d * 20.0d);
        assertTrue(maxRecordCoordinate(aSquared) <= 9.75d * 20.0d);
        assertTrue(maxRecordCoordinate(bSquared) <= 9.75d * 20.0d);
    }

    @Test
    void simpleLinearFormulasUseCompactFontHeightButLeftRightScriptsDoNot() throws IOException {
        byte[] simple = VectorWmfFormulaRenderer.render("S_{1}\\colon S_{3}=a^{2}\\colon b^{2}", 64.0d, 17.25d);
        byte[] shortEquation = VectorWmfFormulaRenderer.render("S_{2}=2", 30.0d, 15.75d);
        byte[] shortSuperscriptEquation = VectorWmfFormulaRenderer.render("a^{2}=1", 30.0d, 15.75d);
        byte[] shortScript = VectorWmfFormulaRenderer.render("S_{2}", 12.0d, 15.75d);
        byte[] leftRight = VectorWmfFormulaRenderer.render("\\left(a+b\\right)^2", 36.0d, 18.75d);
        byte[] geometryLabel = VectorWmfFormulaRenderer.render("\\pwmetrics{17.992,11.995,18.000,12.000}AB", 18.0d,
            12.0d);
        byte[] mixedGeometryText = VectorWmfFormulaRenderer.render("AB1", 18.0d, 12.0d);

        assertTrue(createFontHeights(simple).get(0) > -240);
        assertTrue(createFontHeights(simple).get(0) > createFontHeights(shortScript).get(0));
        assertTrue(createFontHeights(shortEquation).get(0) < createFontHeights(simple).get(0));
        assertTrue(createFontHeights(geometryLabel).get(0) < createFontHeights(mixedGeometryText).get(0));
        assertTrue(VectorWmfFormulaRenderer.isStandaloneTallGeometryLabel(
            "\\pwmetrics{17.992,11.995,18.000,12.000}AB"));
        assertTrue(VectorWmfFormulaRenderer.isStandaloneTallGeometryLabel(
            "\\pwmetrics{17.992,11.995,18.000,12.000}BD"));
        assertFalse(VectorWmfFormulaRenderer.isStandaloneTallGeometryLabel("AC"));
        assertFalse(VectorWmfFormulaRenderer.isStandaloneTallGeometryLabel("CD"));
        assertFalse(VectorWmfFormulaRenderer.isStandaloneTallGeometryLabel("ABCD"));
        assertFalse(VectorWmfFormulaRenderer.isStandaloneTallGeometryLabel("A1"));
        assertEquals(createFontHeights(shortEquation).get(0), createFontHeights(shortSuperscriptEquation).get(0));
        assertEquals(1.0d, VectorWmfFormulaRenderer.scriptRelationFontYScale("S_{2}=2"));
        assertTrue(VectorWmfFormulaRenderer.shortScriptEquationWidthScale("S_{2}=2") < 1.0d);
        assertTrue(VectorWmfFormulaRenderer.shortScriptEquationWidthScale("a^{2}=1") < 1.0d);
        assertEquals(0.969d, VectorWmfFormulaRenderer.shortExactEquationWidthScale("S_{2}=2"));
        assertEquals(1.0d, VectorWmfFormulaRenderer.shortExactEquationWidthScale("S_{2}=2=a\\times b"));
        assertEquals(0.965d, VectorWmfFormulaRenderer.shortExactEquationWidthScale("b=2"));
        assertEquals(1.025d, VectorWmfFormulaRenderer.shortExactEquationWidthScale("a=1"));
        assertEquals(1.025d, VectorWmfFormulaRenderer.shortExactEquationWidthScale(
            "\\pwmetrics{22.980,12.989,23.250,12.750}a=1"));
        assertEquals(1.0d, VectorWmfFormulaRenderer.shortExactEquationWidthScale("a\\colon b=5\\colon 7"));
        assertEquals(1.0d, VectorWmfFormulaRenderer.shortScriptEquationWidthScale(
            "S_{1}\\colon S_{3}=a^{2}\\colon b^{2}"));
        assertTrue(maxTextRightCoordinate(shortEquation) < 568);
        assertTrue(maxTextRightCoordinate(shortSuperscriptEquation) < 30.0d * 20.0d);
        assertEquals(1.0d, VectorWmfFormulaRenderer.shortScriptEquationWidthScale("S_{2}"));
        assertTrue(createFontHeights(simple).get(0) > createFontHeights(shortScript).get(0));
        assertEquals(-245, createFontHeights(leftRight).get(0));
        assertTrue(VectorWmfFormulaRenderer.isShortScriptEquation("S_{2}=2"));
        assertTrue(VectorWmfFormulaRenderer.isShortScriptEquation("S^{2}=4"));
        assertTrue(VectorWmfFormulaRenderer.isShortScriptEquation("a^{2}=1"));
        assertTrue(VectorWmfFormulaRenderer.isShortScriptEquation("a_{2}=1"));
        assertTrue(VectorWmfFormulaRenderer.isShortScriptEquation("x_{1}=2"));
        assertTrue(VectorWmfFormulaRenderer.isShortScriptEquation("A^{2}=1"));
        assertTrue(VectorWmfFormulaRenderer.isShortScriptEquation("S_{2}=2.5"));
        assertFalse(VectorWmfFormulaRenderer.isShortScriptEquation("S_{1}\\colon S_{3}=a^{2}\\colon b^{2}"));
        assertFalse(VectorWmfFormulaRenderer.isShortScriptEquation("S=\\left(a+b\\right)^2=9"));
        assertFalse(VectorWmfFormulaRenderer.canRender("\\dfrac{a}{b}^2"));
        assertFalse(VectorWmfFormulaRenderer.canRender("\\cfrac{a}{b}^2"));
    }

    @Test
    void closingFenceSuperscriptsUseRaisedBaseline() throws IOException {
        byte[] leftRight = VectorWmfFormulaRenderer.render("\\left(a+b\\right)^2", 36.0d, 18.75d);

        assertTrue(VectorWmfFormulaRenderer.canRender("\\left(a+b\\right)^2"));
        assertTrue(minTextYCoordinate(leftRight) <= 115);
        assertTrue(maxRecordCoordinate(leftRight) <= 36.0d * 20.0d);
    }

    @Test
    void symbolCommandsDoNotEncodeAsQuestionMarks() throws IOException {
        byte[] wmf = VectorWmfFormulaRenderer.render(
            "\\times+\\div+\\leq+\\geq+\\neq+\\pi+\\cdots+\\parallel+\\because", 180.0d, 16.0d);
        List<byte[]> textRecords = extTextOutBytes(wmf);

        assertTrue(VectorWmfFormulaRenderer.canRender(
            "\\times+\\div+\\leq+\\geq+\\neq+\\pi+\\cdots+\\parallel+\\because"));
        assertFalse(textRecords.isEmpty());
        assertFalse(textRecords.stream().anyMatch(VectorWmfFormulaRendererTest::containsQuestionMark));
    }

    @Test
    void flatAlignmentMarkersAreNotRenderedAsText() throws IOException {
        byte[] wmf = VectorWmfFormulaRenderer.render("&=8.27\\times (100+3)", 76.0d, 15.0d);
        List<byte[]> textRecords = extTextOutBytes(wmf);

        assertTrue(VectorWmfFormulaRenderer.canRender("&=8.27\\times (100+3)"));
        assertFalse(textRecords.stream().anyMatch(bytes -> containsByte(bytes, (byte) '&')));
        assertTrue(textRecords.stream().anyMatch(bytes -> containsByte(bytes, (byte) '=')));
    }

    @Test
    void escapedAmpersandInsideArrayCellIsNotSplitAsAlignment() throws IOException {
        byte[] wmf = VectorWmfFormulaRenderer.render("\\begin{array}{c}a\\&b\\end{array}", 32.0d, 15.0d);
        List<byte[]> textRecords = extTextOutBytes(wmf);

        assertTrue(VectorWmfFormulaRenderer.canRender("\\begin{array}{c}a\\&b\\end{array}"));
        assertTrue(textRecords.stream().anyMatch(bytes -> containsByte(bytes, (byte) '&')));
    }

    @Test
    void multiplicationAndDivisionUseSymbolGlyphBytesInWmfPreview() throws IOException {
        byte[] wmf = VectorWmfFormulaRenderer.render("8.27\\times 100\\div 2", 92.0d, 15.0d);
        List<byte[]> textRecords = extTextOutBytes(wmf);

        assertFalse(textRecords.isEmpty());
        assertTrue(textRecords.stream().anyMatch(bytes -> containsByte(bytes, (byte) 0xB4)));
        assertTrue(textRecords.stream().anyMatch(bytes -> containsByte(bytes, (byte) 0xB8)));
    }

    @Test
    void rightArrowUsesVectorLineInWmfPreview() throws IOException {
        byte[] wmf = VectorWmfFormulaRenderer.render("2H_{2}+O_{2}\\rightarrow 2H_{2}O", 132.0d, 18.0d);
        List<byte[]> textRecords = extTextOutBytes(wmf);
        List<Integer> records = records(wmf);

        assertTrue(VectorWmfFormulaRenderer.canRender("2H_{2}+O_{2}\\rightarrow 2H_{2}O"));
        assertTrue(records.contains(0x0325));
        assertFalse(textRecords.stream().anyMatch(VectorWmfFormulaRendererTest::containsQuestionMark));
    }

    @Test
    void shortArrowCommandDoesNotMatchLongerCommandPrefix() throws IOException {
        byte[] wmf = VectorWmfFormulaRenderer.render("\\text{to}+A", 42.0d, 15.0d);
        List<Integer> records = records(wmf);

        assertTrue(VectorWmfFormulaRenderer.canRender("\\text{to}+A"));
        assertFalse(records.contains(0x0325));
    }

    @Test
    void latinFormulaLettersUseTimesNewRomanItalicPreviewFont() throws IOException {
        byte[] wmf = VectorWmfFormulaRenderer.render("S_{AOB}=12\\times AB", 96.0d, 18.0d);
        byte[] digits = VectorWmfFormulaRenderer.render("12", 16.0d, 13.0d);
        List<byte[]> textRecords = extTextOutBytes(wmf);

        assertTrue(selectedTextFontItalic(wmf, "S"));
        assertTrue(selectedTextFontItalic(wmf, "AOB"));
        assertTrue(selectedTextFontItalic(wmf, "AB"));
        assertFalse(selectedTextFontItalic(digits, "12"));
        assertTrue(textRecords.stream().anyMatch(bytes -> containsByte(bytes, (byte) 0xB4)));
        assertFalse(selectedTextBytesFontItalic(wmf, new byte[] {(byte) 0xB4}));
    }

    @Test
    void extTextOutKeepsTransparentOptionsAndCarriesDxArray() throws IOException {
        byte[] wmf = VectorWmfFormulaRenderer.render("8.27\\times 100", 72.0d, 15.0d);

        assertTrue(hasExtTextOutDxArray(wmf));
    }

    @Test
    void cjkExtTextOutCarriesByteAlignedDxArray() throws IOException {
        byte[] wmf = VectorWmfFormulaRenderer.render("3\\mathrm{个篮球}", 60.0d, 15.0d);

        assertTrue(hasExtTextOutDxArray(wmf));
    }

    @Test
    void scaledPreviewCoordinatesStayInsideWmfWindow() throws IOException {
        byte[] wmf = VectorWmfFormulaRenderer.render(
            "\\left(\\frac{1}{2}+\\frac{1}{3}+\\cdots+\\frac{1}{20}\\right)", 40.0d, 16.0d);

        assertTrue(maxRecordCoordinate(wmf) <= 40.0d * 20.0d);
    }

    @Test
    void tallArraysKeepHorizontalScaleWhenHeightIsCompressed() throws IOException {
        String latex = "\\begin{array}{ccccccccc} 1 & & & & & & & & \\\\ "
            + "2 & 3 & 4 & & & & & & \\\\ "
            + "5 & 6 & 7 & 8 & 9 & & & & \\\\ "
            + "10 & 11 & 12 & 13 & 14 & 15 & 16 & & \\\\ "
            + "17 & 18 & 19 & 20 & 21 & 22 & 23 & 24 & 25\\\\ "
            + "\\cdots & \\cdots & & & & & & & \\end{array}";
        byte[] wmf = VectorWmfFormulaRenderer.render(latex, 108.0d, 14.0d);

        assertEquals((int) (108.0d * 20.0d), windowExtX(wmf));
        assertEquals((int) (14.0d * 20.0d), windowExtY(wmf));
        assertTrue(maxTextRightCoordinate(wmf) > 90.0d * 20.0d);
        assertTrue(maxRecordCoordinate(wmf) <= 108.0d * 20.0d);
        assertTrue(maxTextYCoordinate(wmf) <= 14.0d * 20.0d);
    }

    private static List<Integer> records(byte[] data) {
        int offset = hasPlaceableHeader(data) ? 22 : 0;
        offset += 18;
        List<Integer> out = new ArrayList<>();
        while (offset + 6 <= data.length) {
            int sizeWords = dword(data, offset);
            int function = word(data, offset + 4);
            out.add(function);
            if (function == 0x0000 || sizeWords <= 0) {
                break;
            }
            offset += sizeWords * 2;
        }
        return out;
    }

    private static List<byte[]> extTextOutBytes(byte[] data) {
        int offset = hasPlaceableHeader(data) ? 22 : 0;
        offset += 18;
        List<byte[]> out = new ArrayList<>();
        while (offset + 6 <= data.length) {
            int sizeWords = dword(data, offset);
            int function = word(data, offset + 4);
            if (function == 0x0000 || sizeWords <= 0) {
                break;
            }
            if (function == 0x0A32 && offset + 14 <= data.length) {
                int count = word(data, offset + 10);
                int textOffset = offset + 14;
                if (count >= 0 && textOffset + count <= data.length) {
                    byte[] text = new byte[count];
                    System.arraycopy(data, textOffset, text, 0, count);
                    out.add(text);
                }
            }
            offset += sizeWords * 2;
        }
        return out;
    }

    private static List<Integer> createFontHeights(byte[] data) {
        int offset = hasPlaceableHeader(data) ? 22 : 0;
        offset += 18;
        List<Integer> out = new ArrayList<>();
        while (offset + 8 <= data.length) {
            int sizeWords = dword(data, offset);
            int function = word(data, offset + 4);
            if (function == 0x0000 || sizeWords <= 0) {
                break;
            }
            if (function == 0x02FB) {
                out.add(signedWord(data, offset + 6));
            }
            offset += sizeWords * 2;
        }
        return out;
    }

    private static List<Integer> createFontWidths(byte[] data) {
        int offset = hasPlaceableHeader(data) ? 22 : 0;
        offset += 18;
        List<Integer> out = new ArrayList<>();
        while (offset + 10 <= data.length) {
            int sizeWords = dword(data, offset);
            int function = word(data, offset + 4);
            if (function == 0x0000 || sizeWords <= 0) {
                break;
            }
            if (function == 0x02FB) {
                out.add(signedWord(data, offset + 8));
            }
            offset += sizeWords * 2;
        }
        return out;
    }

    private static boolean containsQuestionMark(byte[] bytes) {
        return containsByte(bytes, (byte) '?');
    }

    private static boolean hasExtTextOutDxArray(byte[] data) {
        int offset = hasPlaceableHeader(data) ? 22 : 0;
        offset += 18;
        while (offset + 6 <= data.length) {
            int sizeWords = dword(data, offset);
            int function = word(data, offset + 4);
            if (function == 0x0000 || sizeWords <= 0) {
                break;
            }
            if (function == 0x0A32 && offset + 14 <= data.length) {
                int count = word(data, offset + 10);
                int options = word(data, offset + 12);
                int textBytes = count + (count & 1);
                int payloadBytes = sizeWords * 2 - 6;
                if (options == 0 && count > 0 && payloadBytes >= 8 + textBytes + count * 2) {
                    return true;
                }
            }
            offset += sizeWords * 2;
        }
        return false;
    }

    private static boolean containsByte(byte[] bytes, byte target) {
        for (byte value : bytes) {
            if (value == target) {
                return true;
            }
        }
        return false;
    }

    private static int maxRecordCoordinate(byte[] data) {
        int offset = hasPlaceableHeader(data) ? 22 : 0;
        offset += 18;
        int max = 0;
        while (offset + 6 <= data.length) {
            int sizeWords = dword(data, offset);
            int function = word(data, offset + 4);
            if (function == 0x0000 || sizeWords <= 0) {
                break;
            }
            if (function == 0x0A32 && offset + 14 <= data.length) {
                int x = word(data, offset + 8);
                int count = word(data, offset + 10);
                int textBytes = count + (count & 1);
                int dxOffset = offset + 14 + textBytes;
                int dxTotal = 0;
                while (dxOffset + 2 <= offset + sizeWords * 2) {
                    dxTotal += word(data, dxOffset);
                    dxOffset += 2;
                }
                max = Math.max(max, x + dxTotal);
            } else if (function == 0x0325 && offset + 18 <= data.length) {
                max = Math.max(max, word(data, offset + 10));
                max = Math.max(max, word(data, offset + 14));
            }
            offset += sizeWords * 2;
        }
        return max;
    }

    private static int firstTextDxTotal(byte[] data) {
        int offset = hasPlaceableHeader(data) ? 22 : 0;
        offset += 18;
        while (offset + 6 <= data.length) {
            int sizeWords = dword(data, offset);
            int function = word(data, offset + 4);
            if (function == 0x0000 || sizeWords <= 0) {
                break;
            }
            if (function == 0x0A32 && offset + 14 <= data.length) {
                int count = word(data, offset + 10);
                int textBytes = count + (count & 1);
                int dxOffset = offset + 14 + textBytes;
                int dxTotal = 0;
                while (dxOffset + 2 <= offset + sizeWords * 2) {
                    dxTotal += word(data, dxOffset);
                    dxOffset += 2;
                }
                return dxTotal;
            }
            offset += sizeWords * 2;
        }
        return 0;
    }

    private static int textDxTotal(byte[] data, String asciiText) {
        byte[] expected = asciiText.getBytes(java.nio.charset.StandardCharsets.US_ASCII);
        int offset = hasPlaceableHeader(data) ? 22 : 0;
        offset += 18;
        while (offset + 6 <= data.length) {
            int sizeWords = dword(data, offset);
            int function = word(data, offset + 4);
            if (function == 0x0000 || sizeWords <= 0) {
                break;
            }
            if (function == 0x0A32 && offset + 14 <= data.length) {
                int count = word(data, offset + 10);
                int textOffset = offset + 14;
                if (count == expected.length && bytesEqual(data, textOffset, expected)) {
                    int dxOffset = textOffset + count + (count & 1);
                    int dxTotal = 0;
                    while (dxOffset + 2 <= offset + sizeWords * 2) {
                        dxTotal += word(data, dxOffset);
                        dxOffset += 2;
                    }
                    return dxTotal;
                }
            }
            offset += sizeWords * 2;
        }
        return 0;
    }

    private static int textFontHeightTwips(byte[] data, String asciiText) {
        byte[] expected = asciiText.getBytes(java.nio.charset.StandardCharsets.US_ASCII);
        List<Integer> fontHeights = new ArrayList<>();
        int offset = hasPlaceableHeader(data) ? 22 : 0;
        offset += 18;
        int selected = -1;
        while (offset + 6 <= data.length) {
            int sizeWords = dword(data, offset);
            int function = word(data, offset + 4);
            if (function == 0x0000 || sizeWords <= 0) {
                break;
            }
            if (function == 0x02FB && offset + 8 <= data.length) {
                fontHeights.add(Math.abs(signedWord(data, offset + 6)));
            } else if (function == 0x012D && offset + 8 <= data.length) {
                selected = word(data, offset + 6);
            } else if (function == 0x0A32 && offset + 14 <= data.length) {
                int count = word(data, offset + 10);
                int textOffset = offset + 14;
                if (selected >= 0 && selected < fontHeights.size()
                    && count == expected.length && bytesEqual(data, textOffset, expected)) {
                    return fontHeights.get(selected);
                }
            }
            offset += sizeWords * 2;
        }
        return 0;
    }

    private static int maxSelectedFontHeightTwips(byte[] data) {
        List<Integer> fontHeights = new ArrayList<>();
        int offset = hasPlaceableHeader(data) ? 22 : 0;
        offset += 18;
        int selected = -1;
        int max = 0;
        while (offset + 6 <= data.length) {
            int sizeWords = dword(data, offset);
            int function = word(data, offset + 4);
            if (function == 0x0000 || sizeWords <= 0) {
                break;
            }
            if (function == 0x02FB && offset + 8 <= data.length) {
                fontHeights.add(Math.abs(signedWord(data, offset + 6)));
            } else if (function == 0x012D && offset + 8 <= data.length) {
                selected = word(data, offset + 6);
            } else if (function == 0x0A32 && selected >= 0 && selected < fontHeights.size()) {
                max = Math.max(max, fontHeights.get(selected));
            }
            offset += sizeWords * 2;
        }
        return max;
    }

    private static boolean selectedTextFontItalic(byte[] data, String asciiText) {
        return selectedTextBytesFontItalic(data, asciiText.getBytes(java.nio.charset.StandardCharsets.US_ASCII));
    }

    private static boolean selectedTextBytesFontItalic(byte[] data, byte[] expected) {
        List<Boolean> fontItalics = new ArrayList<>();
        int offset = hasPlaceableHeader(data) ? 22 : 0;
        offset += 18;
        int selected = -1;
        while (offset + 6 <= data.length) {
            int sizeWords = dword(data, offset);
            int function = word(data, offset + 4);
            if (function == 0x0000 || sizeWords <= 0) {
                break;
            }
            if (function == 0x02FB && offset + 17 <= data.length) {
                fontItalics.add((data[offset + 16] & 0xff) != 0);
            } else if (function == 0x012D && offset + 8 <= data.length) {
                selected = word(data, offset + 6);
            } else if (function == 0x0A32 && offset + 14 <= data.length) {
                int count = word(data, offset + 10);
                int textOffset = offset + 14;
                if (selected >= 0 && selected < fontItalics.size()
                    && count == expected.length && bytesEqual(data, textOffset, expected)) {
                    return fontItalics.get(selected);
                }
            }
            offset += sizeWords * 2;
        }
        throw new AssertionError("Text record not found: " + java.util.Arrays.toString(expected));
    }

    private static boolean bytesEqual(byte[] data, int offset, byte[] expected) {
        if (offset < 0 || offset + expected.length > data.length) {
            return false;
        }
        for (int i = 0; i < expected.length; i++) {
            if (data[offset + i] != expected[i]) {
                return false;
            }
        }
        return true;
    }

    private static int firstTextAverageDx(byte[] data) {
        int offset = hasPlaceableHeader(data) ? 22 : 0;
        offset += 18;
        while (offset + 6 <= data.length) {
            int sizeWords = dword(data, offset);
            int function = word(data, offset + 4);
            if (function == 0x0000 || sizeWords <= 0) {
                break;
            }
            if (function == 0x0A32 && offset + 14 <= data.length) {
                int count = word(data, offset + 10);
                int textBytes = count + (count & 1);
                int dxOffset = offset + 14 + textBytes;
                int dxTotal = 0;
                int dxCount = 0;
                while (dxOffset + 2 <= offset + sizeWords * 2) {
                    dxTotal += word(data, dxOffset);
                    dxCount++;
                    dxOffset += 2;
                }
                return dxCount == 0 ? 0 : dxTotal / dxCount;
            }
            offset += sizeWords * 2;
        }
        return 0;
    }

    private static int totalTextDx(byte[] data) {
        int offset = hasPlaceableHeader(data) ? 22 : 0;
        offset += 18;
        int total = 0;
        while (offset + 6 <= data.length) {
            int sizeWords = dword(data, offset);
            int function = word(data, offset + 4);
            if (function == 0x0000 || sizeWords <= 0) {
                break;
            }
            if (function == 0x0A32 && offset + 14 <= data.length) {
                int count = word(data, offset + 10);
                int textBytes = count + (count & 1);
                int dxOffset = offset + 14 + textBytes;
                while (dxOffset + 2 <= offset + sizeWords * 2) {
                    total += word(data, dxOffset);
                    dxOffset += 2;
                }
            }
            offset += sizeWords * 2;
        }
        return total;
    }

    private static int minSelectedFontHeightTwips(byte[] data) {
        List<Integer> fontHeights = new ArrayList<>();
        int offset = hasPlaceableHeader(data) ? 22 : 0;
        offset += 18;
        int selected = -1;
        int min = Integer.MAX_VALUE;
        while (offset + 6 <= data.length) {
            int sizeWords = dword(data, offset);
            int function = word(data, offset + 4);
            if (function == 0x0000 || sizeWords <= 0) {
                break;
            }
            if (function == 0x02FB && offset + 8 <= data.length) {
                fontHeights.add(Math.abs(signedWord(data, offset + 6)));
            } else if (function == 0x012D && offset + 8 <= data.length) {
                selected = word(data, offset + 6);
            } else if (function == 0x0A32 && selected >= 0 && selected < fontHeights.size()) {
                min = Math.min(min, fontHeights.get(selected));
            }
            offset += sizeWords * 2;
        }
        return min == Integer.MAX_VALUE ? 0 : min;
    }

    private static int maxTextRightCoordinate(byte[] data) {
        int offset = hasPlaceableHeader(data) ? 22 : 0;
        offset += 18;
        int max = 0;
        while (offset + 6 <= data.length) {
            int sizeWords = dword(data, offset);
            int function = word(data, offset + 4);
            if (function == 0x0000 || sizeWords <= 0) {
                break;
            }
            if (function == 0x0A32 && offset + 14 <= data.length) {
                int x = word(data, offset + 8);
                max = Math.max(max, x + textDxTotalAt(data, offset, sizeWords));
            }
            offset += sizeWords * 2;
        }
        return max;
    }

    private static int maxTextYCoordinate(byte[] data) {
        int offset = hasPlaceableHeader(data) ? 22 : 0;
        offset += 18;
        int max = 0;
        while (offset + 6 <= data.length) {
            int sizeWords = dword(data, offset);
            int function = word(data, offset + 4);
            if (function == 0x0000 || sizeWords <= 0) {
                break;
            }
            if (function == 0x0A32 && offset + 14 <= data.length) {
                max = Math.max(max, word(data, offset + 6));
            }
            offset += sizeWords * 2;
        }
        return max;
    }

    private static int minTextYCoordinateBetween(byte[] data, int minX, int maxX) {
        int offset = hasPlaceableHeader(data) ? 22 : 0;
        offset += 18;
        int min = Integer.MAX_VALUE;
        while (offset + 6 <= data.length) {
            int sizeWords = dword(data, offset);
            int function = word(data, offset + 4);
            if (function == 0x0000 || sizeWords <= 0) {
                break;
            }
            if (function == 0x0A32 && offset + 14 <= data.length) {
                int y = word(data, offset + 6);
                int x = word(data, offset + 8);
                if (x >= minX && x <= maxX) {
                    min = Math.min(min, y);
                }
            }
            offset += sizeWords * 2;
        }
        assertTrue(min != Integer.MAX_VALUE, "expected ExtTextOut records inside x range");
        return min;
    }

    private static int maxTextYCoordinateBetween(byte[] data, int minX, int maxX) {
        int offset = hasPlaceableHeader(data) ? 22 : 0;
        offset += 18;
        int max = Integer.MIN_VALUE;
        while (offset + 6 <= data.length) {
            int sizeWords = dword(data, offset);
            int function = word(data, offset + 4);
            if (function == 0x0000 || sizeWords <= 0) {
                break;
            }
            if (function == 0x0A32 && offset + 14 <= data.length) {
                int y = word(data, offset + 6);
                int x = word(data, offset + 8);
                if (x >= minX && x <= maxX) {
                    max = Math.max(max, y);
                }
            }
            offset += sizeWords * 2;
        }
        assertTrue(max != Integer.MIN_VALUE, "expected ExtTextOut records inside x range");
        return max;
    }

    private static int minTextXCoordinateBetween(byte[] data, int minX, int maxX) {
        int offset = hasPlaceableHeader(data) ? 22 : 0;
        offset += 18;
        int min = Integer.MAX_VALUE;
        while (offset + 6 <= data.length) {
            int sizeWords = dword(data, offset);
            int function = word(data, offset + 4);
            if (function == 0x0000 || sizeWords <= 0) {
                break;
            }
            if (function == 0x0A32 && offset + 14 <= data.length) {
                int x = word(data, offset + 8);
                if (x >= minX && x <= maxX) {
                    min = Math.min(min, x);
                }
            }
            offset += sizeWords * 2;
        }
        assertTrue(min != Integer.MAX_VALUE, "expected ExtTextOut records inside x range");
        return min;
    }

    private static int maxTextRightCoordinateBetween(byte[] data, int minX, int maxX) {
        int offset = hasPlaceableHeader(data) ? 22 : 0;
        offset += 18;
        int max = Integer.MIN_VALUE;
        while (offset + 6 <= data.length) {
            int sizeWords = dword(data, offset);
            int function = word(data, offset + 4);
            if (function == 0x0000 || sizeWords <= 0) {
                break;
            }
            if (function == 0x0A32 && offset + 14 <= data.length) {
                int x = word(data, offset + 8);
                if (x >= minX && x <= maxX) {
                    max = Math.max(max, x + textDxTotalAt(data, offset, sizeWords));
                }
            }
            offset += sizeWords * 2;
        }
        assertTrue(max != Integer.MIN_VALUE, "expected ExtTextOut records inside x range");
        return max;
    }

    private static int textDxTotalAt(byte[] data, int offset, int sizeWords) {
        int count = word(data, offset + 10);
        int textBytes = count + (count & 1);
        int dxOffset = offset + 14 + textBytes;
        int dxTotal = 0;
        while (dxOffset + 2 <= offset + sizeWords * 2) {
            dxTotal += word(data, dxOffset);
            dxOffset += 2;
        }
        return dxTotal;
    }

    private static int maxPolylineYCoordinate(byte[] data) {
        int offset = hasPlaceableHeader(data) ? 22 : 0;
        offset += 18;
        int max = 0;
        while (offset + 6 <= data.length) {
            int sizeWords = dword(data, offset);
            int function = word(data, offset + 4);
            if (function == 0x0000 || sizeWords <= 0) {
                break;
            }
            if (function == 0x0325 && offset + 8 <= data.length) {
                int pointCount = word(data, offset + 6);
                for (int i = 0; i < pointCount && offset + 10 + i * 4 <= data.length; i++) {
                    max = Math.max(max, word(data, offset + 10 + i * 4));
                }
            }
            offset += sizeWords * 2;
        }
        return max;
    }

    private static List<Polyline> polylines(byte[] data) {
        int offset = hasPlaceableHeader(data) ? 22 : 0;
        offset += 18;
        List<Polyline> out = new ArrayList<>();
        while (offset + 18 <= data.length) {
            int sizeWords = dword(data, offset);
            int function = word(data, offset + 4);
            if (function == 0x0000 || sizeWords <= 0) {
                break;
            }
            if (function == 0x0325 && offset + 8 <= data.length) {
                int pointCount = word(data, offset + 6);
                int[] points = new int[pointCount * 2];
                for (int i = 0; i < pointCount && offset + 10 + i * 4 <= data.length; i++) {
                    points[i * 2] = word(data, offset + 8 + i * 4);
                    points[i * 2 + 1] = word(data, offset + 10 + i * 4);
                }
                out.add(new Polyline(points));
            }
            offset += sizeWords * 2;
        }
        return out;
    }

    private static boolean hasLiftedTallRootTurn(Polyline line) {
        int topIndex = line.pointCount() - 2;
        int midIndex = line.pointCount() >= 7 ? topIndex - 2 : topIndex - 1;
        int rootHeightTwips = line.y(midIndex) - line.y(topIndex)
            + (int) Math.round(MathTypeStructureMetrics.SQRT_TALL_BOTTOM_PAD_PT * 20.0d);
        int oldTurnY = line.y(topIndex) + rootHeightTwips
            - (int) Math.round(MathTypeStructureMetrics.SQRT_BOTTOM_PAD_PT * 20.0d);
        return line.y(midIndex) <= oldTurnY - 70
            && rootHeightTwips > (MathTypeStructureMetrics.SQRT_HEIGHT_PT + 4.0d) * 20.0d;
    }

    private static int radicalTopGapTwips(List<Polyline> lines) {
        List<Polyline> radicals = sortedRadicals(lines);
        assertTrue(radicals.size() >= 2, "expected nested radical polylines");
        return radicalTopY(radicals.get(1)) - radicalTopY(radicals.get(0));
    }

    private static List<Polyline> sortedRadicals(List<Polyline> lines) {
        return lines.stream()
            .filter(VectorWmfFormulaRendererTest::isRadicalPolyline)
            .sorted((left, right) -> Integer.compare(left.x1(), right.x1()))
            .toList();
    }

    private static List<Polyline> sortedMainRadicals(List<Polyline> lines) {
        return lines.stream()
            .filter(VectorWmfFormulaRendererTest::isMainRadicalPolyline)
            .sorted((left, right) -> Integer.compare(left.x1(), right.x1()))
            .toList();
    }

    private static boolean isRadicalPolyline(Polyline line) {
        return line.pointCount() >= 4;
    }

    private static boolean isMainRadicalPolyline(Polyline line) {
        return line.pointCount() >= 4
            && line.y(line.pointCount() - 1) == line.y(line.pointCount() - 2)
            && line.x2() - line.x(line.pointCount() - 2) > 40;
    }

    private static boolean hasCompactRootProfile(byte[] data) {
        return polylines(data).stream().anyMatch(line -> line.pointCount() == 3);
    }

    private static int radicalTopY(Polyline line) {
        return line.y(line.pointCount() - 2);
    }

    private static int radicalTopX(Polyline line) {
        return line.x(line.pointCount() - 2);
    }

    private record Polyline(int[] points) {
        int pointCount() {
            return points.length / 2;
        }

        int x1() {
            return points[0];
        }

        int y1() {
            return points[1];
        }

        int x2() {
            return points[points.length - 2];
        }

        int y2() {
            return points[points.length - 1];
        }

        int x(int index) {
            return points[index * 2];
        }

        int y(int index) {
            return points[index * 2 + 1];
        }
    }

    private static void assertCloseTwips(double expectedPt, int actualTwips) {
        int expectedTwips = (int) Math.round(expectedPt * 20.0d);
        assertTrue(Math.abs(actualTwips - expectedTwips) <= 8,
            "expected " + expectedTwips + " twips, got " + actualTwips);
    }

    private static int minTextXCoordinate(byte[] data) {
        int offset = hasPlaceableHeader(data) ? 22 : 0;
        offset += 18;
        int min = Integer.MAX_VALUE;
        while (offset + 6 <= data.length) {
            int sizeWords = dword(data, offset);
            int function = word(data, offset + 4);
            if (function == 0x0000 || sizeWords <= 0) {
                break;
            }
            if (function == 0x0A32 && offset + 14 <= data.length) {
                min = Math.min(min, word(data, offset + 8));
            }
            offset += sizeWords * 2;
        }
        return min == Integer.MAX_VALUE ? 0 : min;
    }

    private static int minTextYCoordinate(byte[] data) {
        int offset = hasPlaceableHeader(data) ? 22 : 0;
        offset += 18;
        int min = Integer.MAX_VALUE;
        while (offset + 6 <= data.length) {
            int sizeWords = dword(data, offset);
            int function = word(data, offset + 4);
            if (function == 0x0000 || sizeWords <= 0) {
                break;
            }
            if (function == 0x0A32 && offset + 14 <= data.length) {
                min = Math.min(min, word(data, offset + 6));
            }
            offset += sizeWords * 2;
        }
        return min == Integer.MAX_VALUE ? 0 : min;
    }

    private static int windowExtX(byte[] data) {
        return windowExt(data, true);
    }

    private static int windowExtY(byte[] data) {
        return windowExt(data, false);
    }

    private static int windowExt(byte[] data, boolean x) {
        int offset = hasPlaceableHeader(data) ? 22 : 0;
        offset += 18;
        while (offset + 10 <= data.length) {
            int sizeWords = dword(data, offset);
            int function = word(data, offset + 4);
            if (function == 0x0000 || sizeWords <= 0) {
                break;
            }
            if (function == 0x020C && offset + 10 <= data.length) {
                int yExt = word(data, offset + 6);
                int xExt = word(data, offset + 8);
                return x ? xExt : yExt;
            }
            offset += sizeWords * 2;
        }
        return -1;
    }

    private static boolean hasPlaceableHeader(byte[] data) {
        return data.length >= 22 && dword(data, 0) == 0x9AC6CDD7;
    }

    private static int word(byte[] data, int offset) {
        return ByteBuffer.wrap(data, offset, 2).order(ByteOrder.LITTLE_ENDIAN).getShort() & 0xffff;
    }

    private static int signedWord(byte[] data, int offset) {
        return ByteBuffer.wrap(data, offset, 2).order(ByteOrder.LITTLE_ENDIAN).getShort();
    }

    private static int dword(byte[] data, int offset) {
        return ByteBuffer.wrap(data, offset, 4).order(ByteOrder.LITTLE_ENDIAN).getInt();
    }
}
