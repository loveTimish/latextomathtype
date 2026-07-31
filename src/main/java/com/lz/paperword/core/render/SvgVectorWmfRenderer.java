package com.lz.paperword.core.render;

import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;
import org.xml.sax.InputSource;

import javax.xml.XMLConstants;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;

/**
 * 把 MathJax(fit) 输出的 SVG 直接转换为<b>真矢量</b> placeable WMF。
 *
 * <p>严格子集：只支持 {@code svg/g/path/rect/defs} 元素、
 * {@code translate/scale/matrix} 变换、{@code M/L/H/V/Q/T/Z} 路径命令；
 * 遇到任何未知元素、命令或绘制属性立即抛出 {@link SvgVectorWmfException}，
 * 由调用方回退到位图路径，绝不静默丢图。
 *
 * <p>纯字节写出，不依赖 AWT，天然 headless 安全。
 *
 * <p>关键 WMF 约定（spike 验证，GDI+ 与 Word 均可加载）：</p>
 * <ul>
 *   <li>标准头 size 字段 = 记录区 WORD 数 + 9（包含标准头自身），numObjects 写实际创建对象数</li>
 *   <li>SetMapMode(MM_ANISOTROPIC) + SetWindowExt(0,0) + SetWindowOrg(H,W)，
 *       y 轴方向由窗口模式处理，坐标变换不做额外翻转</li>
 *   <li>同一 path 的多条轮廓必须放进<b>同一条</b> POLYPOLYGON 记录，
 *       配合 SetPolyFillMode(WINDING) 才能得到字形孔洞</li>
 *   <li>CreatePenIndirect(PS_NULL) + CreateBrushIndirect(纯黑) 后 SelectObject，stock 画刷不可靠</li>
 * </ul>
 */
public final class SvgVectorWmfRenderer {

    /** 渲染不被支持的 SVG 特性时抛出，调用方据此回退位图。 */
    public static final class SvgVectorWmfException extends Exception {
        public SvgVectorWmfException(String message) {
            super(message);
        }
    }

    /** 逻辑坐标单位：每英寸单位数（0.01mm = 2540）。超出 int16 范围时自动降低。 */
    private static final int PREFERRED_UNITS_PER_INCH = 2540;
    private static final int INT16_SAFE = 32000;

    private static final int REC_EOF = 0x0000;
    private static final int REC_SET_MAP_MODE = 0x0103;
    private static final int REC_SET_POLY_FILL_MODE = 0x0106;
    private static final int REC_SELECT_OBJECT = 0x012D;
    private static final int REC_SET_WINDOW_EXT = 0x020B;
    private static final int REC_SET_WINDOW_ORG = 0x020C;
    private static final int REC_DELETE_OBJECT = 0x02F0;
    private static final int REC_CREATE_PEN_INDIRECT = 0x02FA;
    private static final int REC_CREATE_BRUSH_INDIRECT = 0x02FC;
    private static final int REC_POLY_POLYGON = 0x0538;

    private SvgVectorWmfRenderer() {
    }

    /**
     * 渲染 SVG 为 placeable 矢量 WMF。
     *
     * @param svgBytes  MathJax fit 输出的 SVG
     * @param widthPt   目标物理宽度（磅）
     * @param heightPt  目标物理高度（磅）
     * @return placeable WMF 字节
     * @throws SvgVectorWmfException SVG 超出严格子集或结构非法
     */
    public static byte[] render(byte[] svgBytes, double widthPt, double heightPt)
        throws SvgVectorWmfException {
        if (svgBytes == null || svgBytes.length == 0) {
            throw new SvgVectorWmfException("empty SVG input");
        }
        if (widthPt <= 0d || heightPt <= 0d) {
            throw new SvgVectorWmfException("non-positive target size: " + widthPt + "x" + heightPt);
        }
        Element root = parseSvg(svgBytes);
        double[] viewBox = parseViewBox(root);

        double maxDimPt = Math.max(widthPt, heightPt);
        int unitsPerInch = PREFERRED_UNITS_PER_INCH;
        if (maxDimPt / 72.0d * unitsPerInch > INT16_SAFE) {
            unitsPerInch = Math.max(100, (int) Math.floor(INT16_SAFE * 72.0d / maxDimPt));
        }
        int wUnits = Math.max(1, (int) Math.round(widthPt / 72.0d * unitsPerInch));
        int hUnits = Math.max(1, (int) Math.round(heightPt / 72.0d * unitsPerInch));

        double sxu = wUnits / viewBox[2];
        double syu = hUnits / viewBox[3];
        Affine rootTransform = new Affine(sxu, 0d, 0d, syu, -viewBox[0] * sxu, -viewBox[1] * syu);

        List<List<List<double[]>>> groups = new ArrayList<>();
        Affine svgTransform = rootTransform.multiply(parseTransform(root.getAttribute("transform")));
        walk(root, svgTransform, groups);
        if (groups.isEmpty()) {
            throw new SvgVectorWmfException("SVG produced no drawable polygons");
        }

        try {
            return emitWmf(groups, wUnits, hUnits, unitsPerInch);
        } catch (IOException e) {
            throw new SvgVectorWmfException("WMF emit failed: " + e.getMessage());
        }
    }

    // ---------------------------------------------------------------- SVG 解析

    private static Element parseSvg(byte[] svgBytes) throws SvgVectorWmfException {
        try {
            DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
            factory.setNamespaceAware(true);
            factory.setFeature("http://apache.org/xml/features/disallow-doctype-decl", true);
            factory.setFeature("http://xml.org/sax/features/external-general-entities", false);
            factory.setFeature("http://xml.org/sax/features/external-parameter-entities", false);
            factory.setFeature("http://apache.org/xml/features/nonvalidating/load-external-dtd", false);
            factory.setAttribute(XMLConstants.ACCESS_EXTERNAL_DTD, "");
            factory.setAttribute(XMLConstants.ACCESS_EXTERNAL_SCHEMA, "");
            factory.setXIncludeAware(false);
            factory.setExpandEntityReferences(false);
            DocumentBuilder builder = factory.newDocumentBuilder();
            Document doc = builder.parse(new InputSource(new ByteArrayInputStream(svgBytes)));
            Element root = doc.getDocumentElement();
            if (root == null || !"svg".equals(localName(root))) {
                throw new SvgVectorWmfException("root element is not <svg>");
            }
            return root;
        } catch (SvgVectorWmfException e) {
            throw e;
        } catch (Exception e) {
            throw new SvgVectorWmfException("SVG parse failed: " + e.getMessage());
        }
    }

    private static double[] parseViewBox(Element root) throws SvgVectorWmfException {
        String vb = root.getAttribute("viewBox");
        if (vb == null || vb.isBlank()) {
            throw new SvgVectorWmfException("missing viewBox");
        }
        String[] parts = vb.trim().split("[\\s,]+");
        if (parts.length != 4) {
            throw new SvgVectorWmfException("malformed viewBox: " + vb);
        }
        try {
            double[] v = new double[4];
            for (int i = 0; i < 4; i++) {
                v[i] = Double.parseDouble(parts[i]);
            }
            if (v[2] <= 0d || v[3] <= 0d) {
                throw new SvgVectorWmfException("non-positive viewBox size: " + vb);
            }
            return v;
        } catch (NumberFormatException e) {
            throw new SvgVectorWmfException("malformed viewBox number: " + vb);
        }
    }

    private static String localName(Element el) {
        String name = el.getLocalName();
        return name != null ? name : el.getTagName();
    }

    private static void walk(Element el, Affine m, List<List<List<double[]>>> groups)
        throws SvgVectorWmfException {
        NodeList children = el.getChildNodes();
        for (int i = 0; i < children.getLength(); i++) {
            Node node = children.item(i);
            if (node.getNodeType() != Node.ELEMENT_NODE) {
                continue;
            }
            Element child = (Element) node;
            String tag = localName(child);
            Affine childM = m.multiply(parseTransform(child.getAttribute("transform")));
            switch (tag) {
                case "g":
                    walk(child, childM, groups);
                    break;
                case "path":
                    groups.add(renderPath(child, childM));
                    break;
                case "rect":
                    groups.add(renderRect(child, childM));
                    break;
                case "defs":
                case "title":
                case "desc":
                    break;
                default:
                    throw new SvgVectorWmfException("unsupported SVG element: <" + tag + ">");
            }
        }
    }

    private static List<List<double[]>> renderPath(Element el, Affine m) throws SvgVectorWmfException {
        String d = el.getAttribute("d");
        if (d == null || d.isBlank()) {
            // SVG 规范：无 d 的 path 不绘制任何内容（MathJax 会产出占位空 path），跳过即可
            return new ArrayList<>(0);
        }
        String fill = el.getAttribute("fill");
        if (fill != null && !fill.isEmpty()
            && !"currentColor".equals(fill) && !"#000".equals(fill)
            && !"#000000".equals(fill) && !"black".equals(fill)) {
            throw new SvgVectorWmfException("unsupported path fill: " + fill);
        }
        String stroke = el.getAttribute("stroke");
        if (stroke != null && !stroke.isEmpty() && !"none".equals(stroke)) {
            throw new SvgVectorWmfException("unsupported path stroke: " + stroke);
        }
        List<List<double[]>> contours = new PathParser(d).parse();
        List<List<double[]>> flattened = new ArrayList<>(contours.size());
        for (List<double[]> contour : contours) {
            List<double[]> poly = flattenContour(contour, m);
            if (poly.size() >= 2) {
                flattened.add(poly);
            }
        }
        return flattened;
    }

    private static List<List<double[]>> renderRect(Element el, Affine m) throws SvgVectorWmfException {
        double x = attrDouble(el, "x", 0d);
        double y = attrDouble(el, "y", 0d);
        double w = attrDouble(el, "width", Double.NaN);
        double h = attrDouble(el, "height", Double.NaN);
        if (Double.isNaN(w) || Double.isNaN(h) || w <= 0d || h <= 0d) {
            throw new SvgVectorWmfException("rect with missing/invalid size");
        }
        List<double[]> poly = new ArrayList<>(4);
        poly.add(m.apply(x, y));
        poly.add(m.apply(x + w, y));
        poly.add(m.apply(x + w, y + h));
        poly.add(m.apply(x, y + h));
        List<List<double[]>> group = new ArrayList<>(1);
        group.add(poly);
        return group;
    }

    private static double attrDouble(Element el, String name, double fallback) throws SvgVectorWmfException {
        String v = el.getAttribute(name);
        if (v == null || v.isBlank()) {
            return fallback;
        }
        try {
            return Double.parseDouble(v.trim());
        } catch (NumberFormatException e) {
            throw new SvgVectorWmfException("malformed " + name + " attribute: " + v);
        }
    }

    // ---------------------------------------------------------------- 变换

    /** 2D 仿射变换，等价 SVG matrix(a,b,c,d,e,f)。 */
    private static final class Affine {
        final double a, b, c, d, e, f;

        Affine(double a, double b, double c, double d, double e, double f) {
            this.a = a; this.b = b; this.c = c; this.d = d; this.e = e; this.f = f;
        }

        Affine multiply(Affine o) {
            return new Affine(
                a * o.a + c * o.b,
                b * o.a + d * o.b,
                a * o.c + c * o.d,
                b * o.c + d * o.d,
                a * o.e + c * o.f + e,
                b * o.e + d * o.f + f);
        }

        double[] apply(double x, double y) {
            return new double[]{a * x + c * y + e, b * x + d * y + f};
        }
    }

    private static final Affine IDENTITY = new Affine(1d, 0d, 0d, 1d, 0d, 0d);
    private static final java.util.regex.Pattern TRANSFORM_PATTERN =
        java.util.regex.Pattern.compile("(matrix|translate|scale)\\s*\\(([^)]*)\\)");

    private static Affine parseTransform(String s) throws SvgVectorWmfException {
        if (s == null || s.isBlank()) {
            return IDENTITY;
        }
        Affine m = IDENTITY;
        int matchedEnd = 0;
        java.util.regex.Matcher matcher = TRANSFORM_PATTERN.matcher(s);
        while (matcher.find()) {
            if (s.substring(matchedEnd, matcher.start()).trim().length() > 0) {
                throw new SvgVectorWmfException("unsupported transform syntax: " + s);
            }
            matchedEnd = matcher.end();
            String name = matcher.group(1);
            double[] v = parseNumbers(matcher.group(2));
            Affine t;
            switch (name) {
                case "matrix":
                    if (v.length != 6) {
                        throw new SvgVectorWmfException("matrix() needs 6 args: " + s);
                    }
                    t = new Affine(v[0], v[1], v[2], v[3], v[4], v[5]);
                    break;
                case "translate":
                    if (v.length < 1 || v.length > 2) {
                        throw new SvgVectorWmfException("translate() needs 1-2 args: " + s);
                    }
                    t = new Affine(1d, 0d, 0d, 1d, v[0], v.length > 1 ? v[1] : 0d);
                    break;
                case "scale":
                    if (v.length < 1 || v.length > 2) {
                        throw new SvgVectorWmfException("scale() needs 1-2 args: " + s);
                    }
                    double sx = v[0];
                    double sy = v.length > 1 ? v[1] : sx;
                    t = new Affine(sx, 0d, 0d, sy, 0d, 0d);
                    break;
                default:
                    throw new SvgVectorWmfException("unsupported transform: " + name);
            }
            m = m.multiply(t);
        }
        if (s.substring(matchedEnd).trim().length() > 0) {
            throw new SvgVectorWmfException("unsupported transform syntax: " + s);
        }
        return m;
    }

    private static double[] parseNumbers(String s) throws SvgVectorWmfException {
        String trimmed = s.trim();
        if (trimmed.isEmpty()) {
            return new double[0];
        }
        String[] parts = trimmed.split("[\\s,]+");
        double[] v = new double[parts.length];
        try {
            for (int i = 0; i < parts.length; i++) {
                v[i] = Double.parseDouble(parts[i]);
            }
        } catch (NumberFormatException e) {
            throw new SvgVectorWmfException("malformed number list: " + s);
        }
        return v;
    }

    // ---------------------------------------------------------------- 路径解析

    /** 严格子集路径解析器：M/L/H/V/Q/T/Z（含小写相对形式、隐式重复、T 反射）。 */
    private static final class PathParser {
        private final List<Object> tokens;
        private int pos;
        private double x, y;
        private double subX, subY;
        private Double qcx, qcy;
        private List<List<double[]>> contours;
        private List<double[]> current;

        PathParser(String d) throws SvgVectorWmfException {
            tokens = tokenize(d);
        }

        private static List<Object> tokenize(String d) throws SvgVectorWmfException {
            List<Object> out = new ArrayList<>();
            java.util.regex.Matcher m = java.util.regex.Pattern
                .compile("[MmLlHhVvQqTtZz]|[-+]?(?:\\d*\\.\\d+|\\d+\\.?)(?:[eE][-+]?\\d+)?")
                .matcher(d);
            int end = 0;
            while (m.find()) {
                if (!d.substring(end, m.start()).trim().isEmpty()) {
                    throw new SvgVectorWmfException("unsupported path data near: "
                        + d.substring(end, Math.min(end + 12, d.length())));
                }
                end = m.end();
                String tok = m.group();
                if (tok.length() == 1 && Character.isLetter(tok.charAt(0))) {
                    out.add(tok.charAt(0));
                } else {
                    try {
                        out.add(Double.parseDouble(tok));
                    } catch (NumberFormatException e) {
                        throw new SvgVectorWmfException("malformed path number: " + tok);
                    }
                }
            }
            if (!d.substring(end).trim().isEmpty()) {
                throw new SvgVectorWmfException("unsupported path data tail: " + d.substring(end));
            }
            return out;
        }

        List<List<double[]>> parse() throws SvgVectorWmfException {
            contours = new ArrayList<>();
            current = null;
            Character cmd = null;
            while (pos < tokens.size()) {
                if (tokens.get(pos) instanceof Character) {
                    cmd = (Character) tokens.get(pos);
                    pos++;
                }
                if (cmd == null) {
                    throw new SvgVectorWmfException("path data starts without command");
                }
                char c = cmd;
                switch (c) {
                    case 'Z':
                    case 'z':
                        if (current != null) {
                            current.add(new double[]{subX, subY, 1}); // close marker
                        }
                        cmd = null;
                        break;
                    case 'M':
                    case 'm': {
                        double nx = num();
                        double ny = num();
                        if (c == 'm') {
                            nx += x;
                            ny += y;
                        }
                        moveTo(nx, ny);
                        cmd = c == 'M' ? 'L' : 'l';
                        break;
                    }
                    case 'L':
                    case 'l': {
                        requireCurrent(c);
                        double nx = num();
                        double ny = num();
                        if (c == 'l') {
                            nx += x;
                            ny += y;
                        }
                        lineTo(nx, ny);
                        break;
                    }
                    case 'H':
                    case 'h': {
                        requireCurrent(c);
                        double nx = num();
                        if (c == 'h') {
                            nx += x;
                        }
                        lineTo(nx, y);
                        break;
                    }
                    case 'V':
                    case 'v': {
                        requireCurrent(c);
                        double ny = num();
                        if (c == 'v') {
                            ny += y;
                        }
                        lineTo(x, ny);
                        break;
                    }
                    case 'Q':
                    case 'q': {
                        requireCurrent(c);
                        double cx = num();
                        double cy = num();
                        double nx = num();
                        double ny = num();
                        if (c == 'q') {
                            cx += x; cy += y; nx += x; ny += y;
                        }
                        quadTo(cx, cy, nx, ny);
                        break;
                    }
                    case 'T':
                    case 't': {
                        requireCurrent(c);
                        double nx = num();
                        double ny = num();
                        if (c == 't') {
                            nx += x;
                            ny += y;
                        }
                        double cx = qcx == null ? x : 2d * x - qcx;
                        double cy = qcy == null ? y : 2d * y - qcy;
                        quadTo(cx, cy, nx, ny);
                        break;
                    }
                    default:
                        throw new SvgVectorWmfException("unsupported path command: " + c);
                }
            }
            if (current != null) {
                contours.add(current);
                current = null;
            }
            if (contours.isEmpty()) {
                throw new SvgVectorWmfException("path produced no contours");
            }
            return contours;
        }

        private void requireCurrent(char c) throws SvgVectorWmfException {
            if (current == null) {
                throw new SvgVectorWmfException("command " + c + " before initial M");
            }
        }

        private double num() throws SvgVectorWmfException {
            if (pos >= tokens.size() || !(tokens.get(pos) instanceof Double)) {
                throw new SvgVectorWmfException("path command missing numeric argument");
            }
            double v = (Double) tokens.get(pos);
            pos++;
            return v;
        }

        private void moveTo(double nx, double ny) {
            if (current != null) {
                contours.add(current);
            }
            current = new ArrayList<>();
            current.add(new double[]{nx, ny, 0});
            x = nx;
            y = ny;
            subX = nx;
            subY = ny;
            qcx = null;
            qcy = null;
        }

        private void lineTo(double nx, double ny) {
            current.add(new double[]{nx, ny, 0});
            x = nx;
            y = ny;
            qcx = null;
            qcy = null;
        }

        private void quadTo(double cx, double cy, double nx, double ny) {
            current.add(new double[]{cx, cy, 2});
            current.add(new double[]{nx, ny, 0});
            qcx = cx;
            qcy = cy;
            x = nx;
            y = ny;
        }
    }

    /**
     * 把一条轮廓折线化并施加变换。段格式：{x, y, type}，
     * type 0 = 线段端点，type 2 = 二次曲线控制点（后随 type 0 端点），type 1 = 闭合标记。
     */
    private static List<double[]> flattenContour(List<double[]> contour, Affine m) {
        List<double[]> out = new ArrayList<>();
        double[] start = null;
        int i = 0;
        while (i < contour.size()) {
            double[] seg = contour.get(i);
            if (seg[2] == 1) {
                if (start != null && !out.isEmpty()) {
                    double[] last = out.get(out.size() - 1);
                    if (Math.abs(last[0] - start[0]) > 1e-9 || Math.abs(last[1] - start[1]) > 1e-9) {
                        out.add(start);
                    }
                }
                i++;
                continue;
            }
            if (seg[2] == 2 && i + 1 < contour.size()) {
                double[] end = contour.get(i + 1);
                double[] p0 = out.isEmpty() ? m.apply(seg[0], seg[1]) : out.get(out.size() - 1);
                double[] p1 = m.apply(seg[0], seg[1]);
                double[] p2 = m.apply(end[0], end[1]);
                double span = Math.max(
                    Math.max(Math.abs(p2[0] - p0[0]), Math.abs(p2[1] - p0[1])),
                    Math.max(Math.abs(p1[0] - p0[0]), Math.abs(p1[1] - p0[1])));
                int n = (int) Math.min(24, Math.max(4, Math.round(span / 2.0d) + 1));
                for (int k = 1; k <= n; k++) {
                    double t = (double) k / n;
                    double mt = 1d - t;
                    out.add(new double[]{
                        mt * mt * p0[0] + 2d * mt * t * p1[0] + t * t * p2[0],
                        mt * mt * p0[1] + 2d * mt * t * p1[1] + t * t * p2[1]});
                }
                i += 2;
                continue;
            }
            double[] p = m.apply(seg[0], seg[1]);
            if (out.isEmpty()) {
                start = p;
            }
            out.add(p);
            i++;
        }
        return out;
    }

    // ---------------------------------------------------------------- WMF 输出

    private static byte[] emitWmf(List<List<List<double[]>>> groups, int wUnits, int hUnits,
                                  int unitsPerInch) throws IOException, SvgVectorWmfException {
        ByteArrayOutputStream records = new ByteArrayOutputStream(64 * 1024);
        int maxRecordWords = 0;

        maxRecordWords = Math.max(maxRecordWords, writeRecord(records, REC_SET_MAP_MODE, payload(out -> writeWord(out, 8))));
        maxRecordWords = Math.max(maxRecordWords, writeRecord(records, REC_SET_WINDOW_EXT, payload(out -> {
            writeWord(out, 0);
            writeWord(out, 0);
        })));
        final int w = wUnits;
        final int h = hUnits;
        maxRecordWords = Math.max(maxRecordWords, writeRecord(records, REC_SET_WINDOW_ORG, payload(out -> {
            writeWord(out, h); // 参数顺序 (y, x)
            writeWord(out, w);
        })));
        maxRecordWords = Math.max(maxRecordWords, writeRecord(records, REC_SET_POLY_FILL_MODE, payload(out -> writeWord(out, 2))));
        maxRecordWords = Math.max(maxRecordWords, writeRecord(records, REC_CREATE_PEN_INDIRECT, payload(out -> {
            writeWord(out, 5); // PS_NULL
            writeWord(out, 0);
            writeWord(out, 0);
            writeWord(out, 0);
            writeWord(out, 0);
        })));
        maxRecordWords = Math.max(maxRecordWords, writeRecord(records, REC_CREATE_BRUSH_INDIRECT, payload(out -> {
            writeWord(out, 0); // BS_SOLID
            writeWord(out, 0); // color lo
            writeWord(out, 0); // color hi
            writeWord(out, 0); // hatch
        })));
        maxRecordWords = Math.max(maxRecordWords, writeRecord(records, REC_SELECT_OBJECT, payload(out -> writeWord(out, 0))));
        maxRecordWords = Math.max(maxRecordWords, writeRecord(records, REC_SELECT_OBJECT, payload(out -> writeWord(out, 1))));

        int drawableGroups = 0;
        for (List<List<double[]>> group : groups) {
            if (group.isEmpty()) {
                continue;
            }
            ByteArrayOutputStream params = new ByteArrayOutputStream();
            writeWord(params, group.size());
            for (List<double[]> contour : group) {
                if (contour.size() > 0x7FFF) {
                    throw new SvgVectorWmfException("contour too long: " + contour.size());
                }
                writeWord(params, contour.size());
            }
            for (List<double[]> contour : group) {
                for (double[] p : contour) {
                    writeShort(params, clamp16(p[0]));
                    writeShort(params, clamp16(p[1]));
                }
            }
            byte[] payload = params.toByteArray();
            maxRecordWords = Math.max(maxRecordWords, writeRecord(records, REC_POLY_POLYGON, payload));
            drawableGroups++;
        }
        if (drawableGroups == 0) {
            throw new SvgVectorWmfException("no drawable polygon groups after filtering");
        }

        maxRecordWords = Math.max(maxRecordWords, writeRecord(records, REC_DELETE_OBJECT, payload(out -> writeWord(out, 1))));
        maxRecordWords = Math.max(maxRecordWords, writeRecord(records, REC_DELETE_OBJECT, payload(out -> writeWord(out, 0))));
        maxRecordWords = Math.max(maxRecordWords, writeRecord(records, REC_EOF, new byte[0]));

        byte[] recordBytes = records.toByteArray();
        int fileSizeWords = recordBytes.length / 2 + 9; // 含标准头自身

        ByteArrayOutputStream out = new ByteArrayOutputStream(40 + recordBytes.length);
        writePlaceableHeader(out, wUnits, hUnits, unitsPerInch);
        writeWord(out, 1);          // memory metafile
        writeWord(out, 9);          // header size in WORDs
        writeWord(out, 0x0300);     // version
        writeDWord(out, fileSizeWords);
        writeWord(out, 2);          // numObjects（pen + brush）
        writeDWord(out, maxRecordWords);
        writeWord(out, 0);          // numParams
        out.write(recordBytes);
        return out.toByteArray();
    }

    private static int clamp16(double v) {
        if (v > 32767d) {
            return 32767;
        }
        if (v < -32768d) {
            return -32768;
        }
        return (int) Math.round(v);
    }

    private static void writePlaceableHeader(ByteArrayOutputStream out, int right, int bottom,
                                             int inch) throws IOException {
        ByteArrayOutputStream header = new ByteArrayOutputStream(22);
        writeDWord(header, 0x9AC6CDD7L);
        writeWord(header, 0);
        writeShort(header, 0);
        writeShort(header, 0);
        writeShort(header, right);
        writeShort(header, bottom);
        writeWord(header, inch);
        writeDWord(header, 0);
        byte[] prefix = header.toByteArray();
        int checksum = 0;
        for (int i = 0; i < 10; i++) {
            checksum ^= ((prefix[i * 2] & 0xFF) | ((prefix[i * 2 + 1] & 0xFF) << 8));
        }
        out.write(prefix);
        writeWord(out, checksum);
    }

    private interface PayloadWriter {
        void write(ByteArrayOutputStream out) throws IOException;
    }

    private static byte[] payload(PayloadWriter writer) throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        writer.write(out);
        return out.toByteArray();
    }

    private static int writeRecord(ByteArrayOutputStream out, int function, byte[] payload)
        throws IOException {
        int sizeWords = 3 + payload.length / 2;
        writeDWord(out, sizeWords);
        writeWord(out, function);
        out.write(payload);
        return sizeWords;
    }

    private static void writeWord(ByteArrayOutputStream out, int value) {
        out.write(value & 0xFF);
        out.write((value >>> 8) & 0xFF);
    }

    private static void writeShort(ByteArrayOutputStream out, int value) {
        out.write(value & 0xFF);
        out.write((value >> 8) & 0xFF);
    }

    private static void writeDWord(ByteArrayOutputStream out, long value) {
        out.write((int) (value & 0xFF));
        out.write((int) ((value >>> 8) & 0xFF));
        out.write((int) ((value >>> 16) & 0xFF));
        out.write((int) ((value >>> 24) & 0xFF));
    }

    /** 调试辅助：统计 WMF 中是否含位图记录（应为 false）。 */
    public static boolean containsBitmapRecord(byte[] wmf) {
        if (wmf == null || wmf.length < 40) {
            return false;
        }
        int offset = 22 + 18;
        while (offset + 6 <= wmf.length) {
            long sizeWords = ((wmf[offset] & 0xFFL)
                | ((wmf[offset + 1] & 0xFFL) << 8)
                | ((wmf[offset + 2] & 0xFFL) << 16)
                | ((wmf[offset + 3] & 0xFFL) << 24));
            int func = (wmf[offset + 4] & 0xFF) | ((wmf[offset + 5] & 0xFF) << 8);
            if (func == 0x0F43 || func == 0x0F41 || func == 0x0D5B || func == 0x0940) {
                return true;
            }
            if (sizeWords < 3 || offset + sizeWords * 2 > wmf.length) {
                return false;
            }
            if (func == REC_EOF) {
                return false;
            }
            offset += (int) sizeWords * 2;
        }
        return false;
    }
}
