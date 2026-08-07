package com.lz.paperword.core.render;

import org.apache.batik.anim.dom.SAXSVGDocumentFactory;
import org.apache.batik.anim.dom.SVGOMSVGElement;
import org.apache.batik.bridge.BridgeContext;
import org.apache.batik.bridge.DocumentLoader;
import org.apache.batik.bridge.GVTBuilder;
import org.apache.batik.bridge.UserAgentAdapter;
import org.apache.batik.ext.awt.g2d.AbstractGraphics2D;
import org.apache.batik.ext.awt.g2d.GraphicContext;
import org.apache.batik.gvt.GraphicsNode;
import org.apache.batik.util.XMLResourceDescriptor;
import org.w3c.dom.Document;
import org.w3c.dom.NodeList;

import java.awt.AlphaComposite;
import java.awt.Color;
import java.awt.Composite;
import java.awt.Font;
import java.awt.FontMetrics;
import java.awt.Graphics;
import java.awt.Graphics2D;
import java.awt.GraphicsConfiguration;
import java.awt.Image;
import java.awt.Paint;
import java.awt.Shape;
import java.awt.Transparency;
import java.awt.font.FontRenderContext;
import java.awt.font.GlyphVector;
import java.awt.font.TextLayout;
import java.awt.geom.AffineTransform;
import java.awt.geom.Area;
import java.awt.geom.Path2D;
import java.awt.image.BufferedImage;
import java.awt.image.RenderedImage;
import java.awt.image.renderable.RenderableImage;
import java.io.StringReader;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.text.AttributedCharacterIterator;
import java.util.ArrayList;
import java.util.Collections;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Set;

/** Builds a path-only scene by asking Batik to paint the complete SVG graphics tree. */
final class BatikVectorSceneBuilder {

    private BatikVectorSceneBuilder() {
    }

    static VectorScene build(byte[] svgBytes) throws VectorSceneException {
        if (svgBytes == null || svgBytes.length == 0) {
            throw new VectorSceneException("empty SVG input");
        }
        BridgeContext context = null;
        try {
            List<Font> bundledFonts = BundledVectorFonts.ensureRegistered();
            String parser = XMLResourceDescriptor.getXMLParserClassName();
            SAXSVGDocumentFactory factory = new SAXSVGDocumentFactory(parser);
            Document document = factory.createDocument(null,
                new StringReader(new String(svgBytes, StandardCharsets.UTF_8)));
            if (!(document.getDocumentElement() instanceof SVGOMSVGElement root)) {
                throw new VectorSceneException("root element is not an SVG viewport");
            }

            UserAgentAdapter userAgent = new UserAgentAdapter() {
                @Override
                public void checkLoadExternalResource(org.apache.batik.util.ParsedURL resourceURL,
                                                      org.apache.batik.util.ParsedURL docURL) {
                    throw new SecurityException("external SVG resources are forbidden: " + resourceURL);
                }

                @Override
                public void checkLoadScript(String scriptType, org.apache.batik.util.ParsedURL scriptURL,
                                            org.apache.batik.util.ParsedURL docURL) {
                    throw new SecurityException("SVG scripts are forbidden");
                }
            };
            SceneState state = new SceneState();
            validateTextNodes(document, bundledFonts, state);
            DocumentLoader loader = new DocumentLoader(userAgent);
            context = new BridgeContext(userAgent, loader);
            context.setDynamicState(BridgeContext.STATIC);
            GraphicsNode graphicsNode = new GVTBuilder().build(context, document);

            double viewportWidth = root.getWidth().getBaseVal().getValue();
            double viewportHeight = root.getHeight().getBaseVal().getValue();
            if (!Double.isFinite(viewportWidth) || !Double.isFinite(viewportHeight)
                    || viewportWidth <= 0d || viewportHeight <= 0d) {
                throw new VectorSceneException("SVG viewport is empty: " + viewportWidth + "x" + viewportHeight);
            }

            RecordingGraphics2D graphics = new RecordingGraphics2D(state, bundledFonts);
            graphicsNode.paint(graphics);
            graphics.dispose();
            if (state.failure != null) {
                throw state.failure;
            }
            return new VectorScene(List.copyOf(state.shapes), viewportWidth, viewportHeight,
                Collections.unmodifiableSet(new LinkedHashSet<>(state.outlinedCodePoints)));
        } catch (VectorSceneException e) {
            throw e;
        } catch (SecurityException e) {
            throw new VectorSceneException(e.getMessage());
        } catch (Exception e) {
            throw new VectorSceneException("Batik SVG scene build failed: " + e.getMessage(), e);
        } finally {
            if (context != null) {
                context.dispose();
            }
        }
    }

    private static void validateTextNodes(Document document, List<Font> bundledFonts, SceneState state)
        throws VectorSceneException {
        NodeList textNodes = document.getElementsByTagNameNS("*", "text");
        for (int index = 0; index < textNodes.getLength(); index++) {
            String text = textNodes.item(index).getTextContent();
            if (text == null || text.isEmpty()) {
                continue;
            }
            int[] codePoints = text.codePoints().toArray();
            for (int codePoint : codePoints) {
                if (Character.isWhitespace(codePoint)) {
                    continue;
                }
                boolean displayable = false;
                for (Font font : bundledFonts) {
                    if (font.canDisplay(codePoint)) {
                        displayable = true;
                        break;
                    }
                }
                if (!displayable) {
                    throw new VectorSceneException(String.format(
                        "MISSING_GLYPH U+%04X in fixed vector font set", codePoint));
                }
                state.outlinedCodePoints.add(codePoint);
            }
        }
    }

    record PaintedShape(Shape shape, Color color) {
        PaintedShape {
            shape = new Path2D.Double(shape);
        }
    }

    record VectorScene(List<PaintedShape> shapes, double viewportWidth, double viewportHeight,
                       Set<Integer> outlinedCodePoints) {
    }

    static final class VectorSceneException extends Exception {
        VectorSceneException(String message) {
            super(message);
        }

        VectorSceneException(String message, Throwable cause) {
            super(message, cause);
        }
    }

    private static final class SceneState {
        final List<PaintedShape> shapes = new ArrayList<>();
        final Set<Integer> outlinedCodePoints = new LinkedHashSet<>();
        VectorSceneException failure;
    }

    /** Graphics2D implementation that records only device-space filled outlines. */
    private static final class RecordingGraphics2D extends AbstractGraphics2D {
        private static final BufferedImage FONT_SURFACE = new BufferedImage(1, 1, BufferedImage.TYPE_INT_ARGB);
        private static final Graphics2D FONT_GRAPHICS = FONT_SURFACE.createGraphics();

        private final SceneState state;
        private final List<Font> bundledFonts;
        private boolean disposed;

        RecordingGraphics2D(SceneState state, List<Font> bundledFonts) {
            super(true);
            this.state = state;
            this.bundledFonts = bundledFonts;
            this.gc = new GraphicContext();
        }

        private RecordingGraphics2D(RecordingGraphics2D source) {
            super(source);
            this.state = source.state;
            this.bundledFonts = source.bundledFonts;
        }

        @Override
        public Graphics create() {
            return new RecordingGraphics2D(this);
        }

        @Override
        public void draw(Shape shape) {
            requireActive();
            if (shape == null) {
                return;
            }
            capture(getStroke().createStrokedShape(shape));
        }

        @Override
        public void fill(Shape shape) {
            requireActive();
            if (shape != null) {
                capture(shape);
            }
        }

        private void capture(Shape userShape) {
            Paint paint = getPaint();
            if (!(paint instanceof Color color)) {
                fail("unsupported SVG paint: " + (paint == null ? "null" : paint.getClass().getName()));
                return;
            }
            Composite composite = getComposite();
            if (!(composite instanceof AlphaComposite alpha)
                    || alpha.getRule() != AlphaComposite.SRC_OVER
                    || alpha.getAlpha() < 0.999f || color.getAlpha() != 255) {
                fail("transparent or non-SRC_OVER SVG compositing is unsupported");
                return;
            }

            AffineTransform transform = getTransform();
            Shape deviceShape = transform.createTransformedShape(userShape);
            Shape clip = getClip();
            if (clip != null) {
                Area clipped = new Area(deviceShape);
                clipped.intersect(new Area(transform.createTransformedShape(clip)));
                deviceShape = clipped;
            }
            if (!deviceShape.getBounds2D().isEmpty()) {
                state.shapes.add(new PaintedShape(deviceShape,
                    new Color(color.getRed(), color.getGreen(), color.getBlue())));
            }
        }

        @Override
        public void drawGlyphVector(GlyphVector glyphVector, float x, float y) {
            requireActive();
            Font font = glyphVector.getFont();
            int missing = font.getMissingGlyphCode();
            for (int i = 0; i < glyphVector.getNumGlyphs(); i++) {
                if (glyphVector.getGlyphCode(i) == missing) {
                    fail("MISSING_GLYPH in font " + font.getFontName());
                    return;
                }
            }
            capture(glyphVector.getOutline(x, y));
        }

        @Override
        public void drawString(String text, float x, float y) {
            requireActive();
            if (text == null || text.isEmpty()) {
                return;
            }
            Font outlineFont = selectOutlineFont(getFont(), text);
            int missingAt = outlineFont.canDisplayUpTo(text);
            if (missingAt >= 0) {
                int codePoint = text.codePointAt(missingAt);
                fail(String.format("MISSING_GLYPH U+%04X in fixed vector font set", codePoint));
                return;
            }
            text.codePoints().forEach(state.outlinedCodePoints::add);
            GlyphVector glyphs = outlineFont.createGlyphVector(getFontRenderContext(), text);
            drawGlyphVector(glyphs, x, y);
        }

        private Font selectOutlineFont(Font requested, String text) {
            if (requested.canDisplayUpTo(text) < 0) {
                return requested;
            }
            for (Font bundled : bundledFonts) {
                Font candidate = bundled.deriveFont(requested.getStyle(), requested.getSize2D());
                if (candidate.canDisplayUpTo(text) < 0) {
                    return candidate;
                }
            }
            return requested;
        }

        @Override
        public void drawString(AttributedCharacterIterator iterator, float x, float y) {
            requireActive();
            if (iterator == null) {
                return;
            }
            StringBuilder text = new StringBuilder();
            for (char ch = iterator.first(); ch != AttributedCharacterIterator.DONE; ch = iterator.next()) {
                text.append(ch);
            }
            text.codePoints().forEach(state.outlinedCodePoints::add);
            iterator.first();
            new TextLayout(iterator, getFontRenderContext()).draw(this, x, y);
        }

        @Override
        public boolean drawImage(Image image, int x, int y, java.awt.image.ImageObserver observer) {
            fail("raster drawImage is forbidden in vector WMF previews");
            return false;
        }

        @Override
        public boolean drawImage(Image image, int x, int y, int width, int height,
                                 java.awt.image.ImageObserver observer) {
            fail("raster drawImage is forbidden in vector WMF previews");
            return false;
        }

        @Override
        public void drawRenderedImage(RenderedImage image, AffineTransform transform) {
            fail("raster drawRenderedImage is forbidden in vector WMF previews");
        }

        @Override
        public void drawRenderableImage(RenderableImage image, AffineTransform transform) {
            fail("raster drawRenderableImage is forbidden in vector WMF previews");
        }

        @Override
        public GraphicsConfiguration getDeviceConfiguration() {
            return FONT_GRAPHICS.getDeviceConfiguration();
        }

        @Override
        public FontMetrics getFontMetrics(Font font) {
            return FONT_GRAPHICS.getFontMetrics(font);
        }

        @Override
        public FontRenderContext getFontRenderContext() {
            return FONT_GRAPHICS.getFontRenderContext();
        }

        @Override
        public void setXORMode(Color color) {
            fail("XOR drawing is unsupported in vector WMF previews");
        }

        @Override
        public void copyArea(int x, int y, int width, int height, int dx, int dy) {
            fail("copyArea is unsupported in vector WMF previews");
        }

        @Override
        public void dispose() {
            disposed = true;
        }

        private void requireActive() {
            if (disposed) {
                throw new IllegalStateException("Graphics2D is disposed");
            }
            if (state.failure != null) {
                throw new VectorSceneRuntimeException(state.failure);
            }
        }

        private void fail(String message) {
            VectorSceneException failure = new VectorSceneException(message);
            state.failure = failure;
            throw new VectorSceneRuntimeException(failure);
        }
    }

    private static final class VectorSceneRuntimeException extends RuntimeException {
        VectorSceneRuntimeException(VectorSceneException cause) {
            super(cause);
        }
    }
}
