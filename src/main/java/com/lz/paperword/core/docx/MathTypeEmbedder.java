package com.lz.paperword.core.docx;

import com.lz.paperword.core.latex.LaTeXNode;
import com.lz.paperword.core.latex.LaTeXParser;
import com.lz.paperword.core.mtef.MtefWriter;
import com.lz.paperword.core.ole.OlePackager;
import com.lz.paperword.core.render.LaTeXImageRenderer;
import org.apache.poi.openxml4j.opc.*;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.concurrent.atomic.AtomicInteger;

/**
 * MathType OLE 公式嵌入器 — 将 LaTeX 公式作为可编辑的 MathType OLE 对象嵌入 Word 文档。
 *
 * <h2>嵌入流程概览</h2>
 * <p>本类协调整个嵌入管线，将一个 LaTeX 公式转变为 Word 文档中可双击编辑的 MathType 公式：</p>
 * <ol>
 *   <li><b>MTEF 生成</b> — 通过 {@link MtefWriter} 将 LaTeX AST 转为 MTEF v5 二进制数据
 *       （MathType 的原生公式格式）</li>
 *   <li><b>OLE 打包</b> — 通过 {@link OlePackager} 将 MTEF 数据封装为 OLE2 复合文档
 *       （oleObjectN.bin），包含 MathType CLSID 和 ProgID "Equation.DSMT4"</li>
 *   <li><b>预览渲染</b> — 通过 {@link LaTeXImageRenderer} 将 LaTeX 渲染为 PNG 预览图
 *       （image_eqN.png），作为公式在文档中的静态显示</li>
 *   <li><b>文档组装</b> — 将 OLE 二进制和预览图作为 OPC 包部件添加到 .docx 中，
 *       建立关系引用，并在段落的 run 中写入 VML shape XML</li>
 * </ol>
 *
 * <h2>OPC 包结构</h2>
 * <p>每个嵌入公式在 .docx（实质为 ZIP/OPC 包）中创建以下部件：</p>
 * <ul>
 *   <li>{@code /word/embeddings/oleObjectN.bin} — OLE2 复合文档（MathType 公式数据）</li>
 *   <li>{@code /word/media/image_eqN.png} — PNG 预览图（文档中的静态显示）</li>
 * </ul>
 * <p>并从 document.xml 部件建立两个关系（Relationship）指向上述部件。</p>
 *
 * <h2>VML 形状与 OLE 对象</h2>
 * <p>在 document.xml 的 {@code <w:object>} 元素中生成：</p>
 * <ul>
 *   <li>{@code <v:shapetype>} — 定义图片形状模板（#_x0000_t75）</li>
 *   <li>{@code <v:shape>} — 引用预览图，设置显示尺寸（style:width/height）</li>
 *   <li>{@code <o:OLEObject>} — 引用 OLE 二进制，ProgID="Equation.DSMT4"</li>
 * </ul>
 * <p>用户在 Word 中看到的是预览图；双击预览图时，Word 通过 OLEObject 的关系引用
 * 找到 oleObjectN.bin，读取其中的 CLSID，激活 MathType 编辑器打开公式。</p>
 *
 * @see OlePackager   OLE2 复合文档打包
 * @see LaTeXImageRenderer  LaTeX 公式预览图渲染
 * @see MtefWriter    LaTeX AST → MTEF 二进制转换
 */
public class MathTypeEmbedder {

    private static final Logger log = LoggerFactory.getLogger(MathTypeEmbedder.class);

    /**
     * 像素到磅的转换系数：96 DPI 下 1 像素 = 0.75 磅。
     * 用于将预览图的像素尺寸转换为 VML shape 的 pt 尺寸。
     */
    private static final double PT_PER_PX = 0.75d;

    /** MTEF 写入器：将 LaTeX AST 转为 MTEF v5 二进制数据 */
    private final MtefWriter mtefWriter = new MtefWriter();

    /** OLE 打包器：将 MTEF 数据封装为 OLE2 复合文档 */
    private final OlePackager olePackager = new OlePackager();

    /** 图片渲染器：将 LaTeX 公式渲染为 PNG/EMF 预览图 */
    private final LaTeXImageRenderer imageRenderer = new LaTeXImageRenderer();

    /** OLE 对象序号计数器（线程安全），确保每个公式在文档中有唯一标识 */
    private final AtomicInteger oleCounter = new AtomicInteger(1);

    /**
     * 将 MathType 公式嵌入 Word 段落的指定 run 中。
     *
     * <p>完整的嵌入管线包含 5 个步骤：</p>
     * <ol>
     *   <li><b>MTEF 生成</b> — 将 LaTeX AST 序列化为 MTEF v5 二进制数据</li>
     *   <li><b>OLE 打包</b> — 将 MTEF 数据封装为 OLE2 复合文档（包含 CLSID + ProgID）</li>
     *   <li><b>预览渲染</b> — 将原始 LaTeX 渲染为高分辨率 PNG 预览图</li>
     *   <li><b>OPC 包写入</b> — 将 OLE 二进制和预览图作为部件写入 .docx 包，
     *       并从 document.xml 建立关系引用</li>
     *   <li><b>VML XML 注入</b> — 在 run 中写入 {@code <w:object>} XML，包含
     *       VML 形状（引用预览图）和 OLEObject（引用 OLE 二进制）</li>
     * </ol>
     *
     * <p>如果预览渲染失败，退化为纯文本显示（[公式源码]）。</p>
     *
     * @param paragraph 目标段落
     * @param run       目标 run（将被替换为 OLE 对象 XML）
     * @param latexAst  解析后的 LaTeX 抽象语法树
     * @param rawLatex  原始 LaTeX 字符串（用于图片渲染和错误回退显示）
     */
    public void embedEquation(XWPFParagraph paragraph, XWPFRun run, LaTeXNode latexAst, String rawLatex) {
        try {
            // 步骤 1：将 LaTeX AST 转为 MTEF v5 二进制数据
            // MTEF 是 MathType 的原生公式格式，包含完整的公式树结构
            byte[] mtefData = mtefWriter.write(latexAst);

            // 步骤 2：将 MTEF 数据打包为 OLE2 复合文档
            // 生成的 oleData 包含 4 个 OLE 流 + MathType CLSID
            byte[] oleData = olePackager.packageOle(mtefData);

            // 步骤 3：渲染高分辨率 PNG 预览图（3 倍渲染，报告 1 倍尺寸）
            LaTeXImageRenderer.PreviewImage preview = imageRenderer.renderForOlePreview(rawLatex);
            if (preview == null || preview.data() == null || preview.data().length == 0) {
                // 预览渲染失败：退化为纯文本显示
                log.warn("Could not render preview for: {}", rawLatex);
                run.setText("[" + rawLatex + "]");
                return;
            }

            // 步骤 4：将 OLE 二进制和预览图写入 .docx 的 OPC 包
            OPCPackage pkg = paragraph.getDocument().getPackage();
            int idx = oleCounter.getAndIncrement();  // 获取唯一序号

            // 4a. 创建 OLE 二进制部件：/word/embeddings/oleObjectN.bin
            // Content-Type 为 OLE 对象的标准 MIME 类型
            PackagePartName olePartName = PackagingURIHelper.createPartName(
                "/word/embeddings/oleObject" + idx + ".bin");
            PackagePart olePart = pkg.createPart(olePartName,
                "application/vnd.openxmlformats-officedocument.oleObject");
            try (var os = olePart.getOutputStream()) {
                os.write(oleData);
            }

            // 4b. 创建预览图部件：/word/media/image_eqN.png
            PackagePartName imgPartName = PackagingURIHelper.createPartName(
                "/word/media/image_eq" + idx + "." + preview.extension());
            PackagePart imgPart = pkg.createPart(imgPartName, preview.contentType());
            try (var os = imgPart.getOutputStream()) {
                os.write(preview.data());
            }

            // 4c. 从 document.xml 部件建立关系引用
            // oleRel：指向 OLE 二进制（o:OLEObject 的 r:id 属性引用此关系）
            // imgRel：指向预览图（v:imagedata 的 r:id 属性引用此关系）
            PackagePart docPart = paragraph.getDocument().getPackagePart();
            PackageRelationship oleRel = docPart.addRelationship(
                olePartName, TargetMode.INTERNAL,
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject");
            PackageRelationship imgRel = docPart.addRelationship(
                imgPartName, TargetMode.INTERNAL,
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image");

            // 步骤 5：在 run 中注入 w:object VML XML
            insertOleObjectXml(run, oleRel.getId(), imgRel.getId(), idx, preview.widthPx(), preview.heightPx());

        } catch (Exception e) {
            // 任何步骤失败：退化为纯文本显示，不影响文档其余部分的生成
            log.error("Failed to embed MathType equation: {}", rawLatex, e);
            run.setText("[" + rawLatex + "]");
        }
    }

    /**
     * 在 Word run 中注入包含 VML 形状和 OLE 引用的 {@code <w:object>} XML。
     *
     * <p>使用 {@code ctr.set()} 完全替换 run 的 XML 内容，插入如下结构：</p>
     * <pre>{@code
     * <w:r>
     *   <w:object>
     *     <v:shapetype .../> — 图片形状模板定义（#_x0000_t75，标准图片形状）
     *     <v:shape ...>      — 引用预览图，定义显示尺寸
     *       <v:imagedata r:id="imgRelId" />
     *     </v:shape>
     *     <o:OLEObject ...   — OLE 对象引用
     *       Type="Embed"     — 嵌入式（非链接）
     *       ProgID="Equation.DSMT4"  — MathType 程序标识
     *       r:id="oleRelId"  — 指向 oleObjectN.bin 的关系
     *     />
     *   </w:object>
     * </w:r>
     * }</pre>
     *
     * <p><b>交互机制：</b>用户在 Word 中看到的是 v:shape 引用的预览图（PNG）。
     * 双击预览图时，Word 读取对应的 o:OLEObject，通过 r:id 找到 OLE 二进制部件，
     * 读取其 CLSID，激活 MathType 编辑器，加载 "Equation Native" 流中的 MTEF 数据。
     * MathType 打开后会自动用其内部生成的 WMF 矢量图替换我们的 PNG 预览。</p>
     *
     * @param run       目标 run（内容将被完全替换）
     * @param oleRelId  OLE 对象的关系 ID（r:id），指向 oleObjectN.bin
     * @param imgRelId  预览图的关系 ID（r:id），指向 image_eqN.png
     * @param shapeIdx  形状序号，用于生成唯一的 ShapeID 和 ObjectID
     * @param widthPx   预览图显示宽度（像素）
     * @param heightPx  预览图显示高度（像素）
     */
    private void insertOleObjectXml(XWPFRun run, String oleRelId, String imgRelId,
                                     int shapeIdx, int widthPx, int heightPx) {
        try {
            CTR ctr = run.getCTR();

            // 确保最小尺寸，防止 Word 中出现不可见的零尺寸形状
            widthPx = Math.max(widthPx, 10);
            heightPx = Math.max(heightPx, 10);

            // 构建唯一标识符
            // ShapeID：VML 形状唯一标识，格式为 _x0000_iNNNN（从 1025 起）
            String shapeId = "_x0000_i" + (1024 + shapeIdx);

            // 将像素尺寸转换为磅（pt）用于 VML style 属性
            // 96 DPI 下：1 px = 0.75 pt
            double targetWidthPt = widthPx * PT_PER_PX;
            double targetHeightPt = heightPx * PT_PER_PX;
            String styleWidth = String.format("%.1fpt", targetWidthPt);
            String styleHeight = String.format("%.1fpt", targetHeightPt);

            // ObjectID：OLE 对象在文档中的唯一标识
            // 使用序号而非时间戳，避免同一毫秒内的重复
            String objectId = "_MT_OBJ_" + shapeIdx;

            // 构建完整的 w:r XML，包含命名空间声明和嵌入的 w:object
            // 此 XML 将完全替换 run 的当前内容
            String runXml = "<w:r " +
                "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" " +
                "xmlns:v=\"urn:schemas-microsoft-com:vml\" " +
                "xmlns:o=\"urn:schemas-microsoft-com:office:office\" " +
                "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
                "<w:object>" +

                // === v:shapetype === 图片形状模板定义
                // id="_x0000_t75" 是 Word 中标准图片形状的预定义类型编号
                // coordsize="21600,21600" 定义内部坐标系（EMU 单位）
                // o:spt="75" 表示预设形状类型 75（图片占位符）
                // 内部的 v:formulas 定义了形状的几何计算公式，用于自适应缩放
                "<v:shapetype id=\"_x0000_t75\" coordsize=\"21600,21600\" o:spt=\"75\" o:preferrelative=\"t\" " +
                "path=\"m@4@5l@4@11@9@11@9@5xe\" filled=\"f\" stroked=\"f\">" +
                "<v:stroke joinstyle=\"miter\"/>" +
                "<v:formulas>" +
                "<v:f eqn=\"if lineDrawn pixelLineWidth 0\"/>" +
                "<v:f eqn=\"sum @0 1 0\"/>" +
                "<v:f eqn=\"sum 0 0 @1\"/>" +
                "<v:f eqn=\"prod @2 1 2\"/>" +
                "<v:f eqn=\"prod @3 21600 pixelWidth\"/>" +
                "<v:f eqn=\"prod @3 21600 pixelHeight\"/>" +
                "<v:f eqn=\"sum @0 0 1\"/>" +
                "<v:f eqn=\"prod @6 1 2\"/>" +
                "<v:f eqn=\"prod @7 21600 pixelWidth\"/>" +
                "<v:f eqn=\"sum @8 21600 0\"/>" +
                "<v:f eqn=\"prod @7 21600 pixelHeight\"/>" +
                "<v:f eqn=\"sum @10 21600 0\"/>" +
                "</v:formulas>" +
                "<v:path o:extrusionok=\"f\" gradientshapeok=\"t\" o:connecttype=\"rect\"/>" +
                "<o:lock v:ext=\"edit\" aspectratio=\"t\"/>" +  // 锁定纵横比
                "</v:shapetype>" +

                // === v:shape === 形状实例，引用预览图
                // type="#_x0000_t75" 引用上面定义的 shapetype
                // style 设置显示尺寸（pt 单位）
                // o:ole="" 标记此形状关联了一个 OLE 对象
                // v:imagedata 通过 r:id 引用预览 PNG 图片
                "<v:shape id=\"" + shapeId + "\" type=\"#_x0000_t75\" " +
                "style=\"width:" + styleWidth + ";height:" + styleHeight + "\" o:ole=\"\">" +
                "<v:imagedata r:id=\"" + imgRelId + "\" o:title=\"\"/>" +
                "</v:shape>" +

                // === o:OLEObject === OLE 对象引用
                // Type="Embed" — 嵌入式对象（数据存储在文档包内）
                // ProgID="Equation.DSMT4" — MathType 的程序标识符
                // ShapeID 关联到上面的 v:shape（双击时激活 OLE）
                // DrawAspect="Content" — 以内容形式显示（非图标）
                // r:id 指向 /word/embeddings/oleObjectN.bin
                "<o:OLEObject Type=\"Embed\" ProgID=\"Equation.DSMT4\" " +
                "ShapeID=\"" + shapeId + "\" DrawAspect=\"Content\" " +
                "ObjectID=\"" + objectId + "\" " +
                "r:id=\"" + oleRelId + "\" />" +
                "</w:object></w:r>";

            // 使用 XmlBeans 解析 XML 字符串并完全替换 CTR（run）的内容
            ctr.set(org.apache.xmlbeans.XmlObject.Factory.parse(runXml));

        } catch (Exception e) {
            log.error("Failed to insert OLE XML into run", e);
        }
    }
}
