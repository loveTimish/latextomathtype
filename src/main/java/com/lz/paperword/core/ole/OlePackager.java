package com.lz.paperword.core.ole;

import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.poifs.filesystem.DirectoryNode;
import org.apache.poi.poifs.filesystem.DocumentEntry;
import org.apache.poi.poifs.filesystem.DocumentInputStream;
import org.apache.poi.poifs.filesystem.Entry;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.ByteBuffer;
import java.nio.ByteOrder;
import java.nio.charset.StandardCharsets;

/**
 * 将 MTEF 二进制数据打包为 MathType OLE2 复合文档（Microsoft Structured Storage 格式）。
 *
 * <h2>OLE2 复合文档概述</h2>
 * <p>OLE2 复合文档是微软定义的一种结构化存储格式，类似于文件系统中的"文件内文件系统"。
 * 它允许在一个二进制文件中包含多个"流"（Stream），每个流存储不同类型的数据。
 * MathType 的 OLE 对象需要以下 4 个流：</p>
 * <ul>
 *   <li><b>\001Ole</b> — 20 字节的标准 OLE 嵌入对象头，标识该对象为嵌入式（非链接式）</li>
 *   <li><b>\001CompObj</b> — COM 对象标识流，包含 ProgID "Equation.DSMT4"（MathType 的程序标识符），
 *       Word 通过此流识别该 OLE 对象的类型并关联到 MathType 编辑器</li>
 *   <li><b>Equation Native</b> — 核心数据流，包含 28 字节的 EQNOLEFILEHDR 头部 + MTEF 二进制公式数据。
 *       MTEF（MathType Equation Format）是 MathType 的原生公式描述格式</li>
 *   <li><b>\003ObjInfo</b> — 6 字节的对象显示信息，标识为嵌入式对象</li>
 * </ul>
 *
 * <h2>打包策略</h2>
 * <p>优先使用模板路径（{@link #packageByTemplate}）：加载一个真实的 MathType OLE 模板文件，
 * 仅替换其中的 "Equation Native" 流。这样可以保留模板中难以从零构造的辅助流数据。
 * 如果模板不可用，则退化为手动构建全部 4 个流的方式。</p>
 *
 * <h2>CLSID 标识</h2>
 * <p>根存储节点的 CLSID 设置为 MathType 的标准类标识符
 * {0002CE03-0000-0000-C000-000000000046}，
 * Word 通过此 CLSID 将嵌入的二进制对象与 MathType 关联。</p>
 *
 * <p>生成的 OLE 对象可嵌入 .docx 文件中，被 Word 和 MathType 识别为可编辑公式，
 * 用户双击即可打开 MathType 编辑器进行修改。</p>
 *
 * @see <a href="https://docs.microsoft.com/en-us/openspecs/windows_protocols/ms-oleds/">MS-OLEDS: OLE Data Structures</a>
 */
public class OlePackager {

    private static final Logger log = LoggerFactory.getLogger(OlePackager.class);

    /** 类路径下的 MathType OLE 模板文件路径，用于优先打包策略 */
    private static final String TEMPLATE_RESOURCE = "/mathtype-template/oleObject-template.bin";

    /**
     * MathType 的 CLSID（类标识符）：{0002CE03-0000-0000-C000-000000000046}
     *
     * <p>CLSID 是 COM/OLE 体系中唯一标识某个 COM 类的 128 位 GUID。
     * Windows 注册表中 MathType 注册了此 CLSID，Word 在打开 .docx 时读取
     * OLE 根存储的 CLSID，查找注册表中对应的 COM 服务器（即 MathType），
     * 从而知道如何激活和编辑该嵌入对象。</p>
     *
     * <p>字节序按照微软 GUID 的混合端序规则排列：
     * Data1(4字节小端) + Data2(2字节小端) + Data3(2字节小端) + Data4(8字节大端)</p>
     */
    private static final byte[] MATHTYPE_CLSID = {
        0x03, (byte) 0xCE, 0x02, 0x00,   // Data1: 0x0002CE03（小端序）
        0x00, 0x00,                       // Data2: 0x0000（小端序）
        0x00, 0x00,                       // Data3: 0x0000（小端序）
        (byte) 0xC0, 0x00,               // Data4[0..1]: 0xC000（大端序）
        0x00, 0x00, 0x00, 0x00, 0x00, 0x46  // Data4[2..7]: 000000000046（大端序）
    };

    /**
     * 将 MTEF 二进制数据打包为 MathType OLE2 复合文档。
     *
     * <p>打包流程：
     * <ol>
     *   <li>优先尝试基于模板打包（{@link #packageByTemplate}），复用真实 MathType 模板中的辅助流</li>
     *   <li>若模板不可用，则手动构建全部 4 个 OLE 流并组装为复合文档</li>
     * </ol>
     *
     * <p>手动构建时创建的 4 个流：
     * <ul>
     *   <li>\001Ole — OLE 嵌入对象标准头</li>
     *   <li>\001CompObj — COM 对象标识（ProgID = "Equation.DSMT4"）</li>
     *   <li>Equation Native — EQNOLEFILEHDR + MTEF 公式数据</li>
     *   <li>\003ObjInfo — 嵌入对象显示属性</li>
     * </ul>
     *
     * @param mtefData MTEF v5 格式的二进制公式数据
     * @return 完整的 OLE2 复合文档字节数组，可直接作为 oleObjectN.bin 嵌入 .docx
     * @throws IOException 写入 POIFS 文件系统失败时抛出
     */
    public byte[] packageOle(byte[] mtefData) throws IOException {
        // 优先尝试模板打包路径：复用真实 MathType OLE 模板，仅替换公式数据流
        byte[] packagedByTemplate = packageByTemplate(mtefData);
        if (packagedByTemplate != null) {
            return packagedByTemplate;
        }

        // 模板不可用时的退化路径：从零构建全部 OLE 流
        try (POIFSFileSystem fs = new POIFSFileSystem()) {
            // 1. 写入 "\001Ole" 流 — 20 字节的标准 OLE 嵌入对象头
            fs.createDocument(new ByteArrayInputStream(createOleStream()),
                "\u0001Ole");

            // 2. 写入 "\001CompObj" 流 — COM 对象标识，包含 ProgID "Equation.DSMT4"
            //    Word 根据此流判断嵌入对象类型并关联 MathType 编辑器
            fs.createDocument(new ByteArrayInputStream(createCompObjStream()),
                "\u0001CompObj");

            // 3. 写入 "Equation Native" 流 — 28 字节 EQNOLEFILEHDR 头 + MTEF 二进制数据
            //    这是 MathType 公式的核心数据，包含公式的完整结构描述
            fs.createDocument(new ByteArrayInputStream(createEquationNativeStream(mtefData)),
                "Equation Native");

            // 4. 写入 "\003ObjInfo" 流 — 6 字节对象信息，标记为嵌入式对象（非链接）
            fs.createDocument(new ByteArrayInputStream(createObjInfoStream()),
                "\u0003ObjInfo");

            // 设置根存储节点的 CLSID 为 MathType 的类标识符
            // Word 通过此 CLSID 在系统注册表中查找 MathType COM 服务器
            fs.getRoot().setStorageClsid(
                new org.apache.poi.hpsf.ClassID(MATHTYPE_CLSID, 0));

            // 将 POIFS 文件系统序列化为字节数组
            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            fs.writeFilesystem(baos);
            return baos.toByteArray();
        }
    }

    /**
     * 优先打包路径：加载真实的 MathType OLE 模板，仅替换 "Equation Native" 流。
     *
     * <p>使用模板的优势：
     * <ul>
     *   <li>保留模板中 MathType 生成的辅助流（如 \001Ole、\001CompObj 等），
     *       这些流包含难以从零精确重建的字节序列</li>
     *   <li>确保与 MathType 编辑器的最大兼容性</li>
     * </ul>
     *
     * <p>关键操作：
     * <ol>
     *   <li>读取模板中原有的 "Equation Native" 流头部，用于保留原始 EQNOLEFILEHDR 头字节</li>
     *   <li>删除原 "Equation Native" 流，写入新的（模板头部 + 新 MTEF 数据）</li>
     *   <li>删除 \002OlePres000 流 — 这是模板中缓存的旧公式预览图（WMF 格式），
     *       MathType 7 的 OLE 对象不包含此流。如果保留，Word 会显示模板原始公式
     *       （简单的 "A"）的预览，而非新公式的预览，导致显示内容不一致</li>
     * </ol>
     *
     * @param mtefData MTEF 二进制公式数据
     * @return 打包后的 OLE2 字节数组；模板不可用或打包失败时返回 null
     */
    private byte[] packageByTemplate(byte[] mtefData) {
        try (InputStream in = OlePackager.class.getResourceAsStream(TEMPLATE_RESOURCE)) {
            if (in == null) {
                // 模板资源不存在，返回 null 触发退化路径
                return null;
            }

            byte[] templateBytes = in.readAllBytes();
            try (POIFSFileSystem fs = new POIFSFileSystem(new ByteArrayInputStream(templateBytes))) {
                DirectoryNode root = fs.getRoot();

                // 读取模板中原有的 "Equation Native" 流，提取其 EQNOLEFILEHDR 头部
                // 以便在构建新流时复用模板头部的字段值
                byte[] templateEquationNative = null;
                if (root.hasEntry("Equation Native")) {
                    Entry nativeEntry = root.getEntry("Equation Native");
                    if (nativeEntry instanceof DocumentEntry documentEntry) {
                        try (DocumentInputStream dis = new DocumentInputStream(documentEntry)) {
                            templateEquationNative = dis.readAllBytes();
                        }
                    }
                    // 删除旧的 "Equation Native" 流，稍后写入替换后的新流
                    nativeEntry.delete();
                }

                // 创建新的 "Equation Native" 流：保留模板头部 + 替换 MTEF 数据负载
                root.createDocument("Equation Native",
                    new ByteArrayInputStream(createEquationNativeStreamFromTemplate(templateEquationNative, mtefData)));

                // 删除过期的 \002OlePres000 预览缓存流
                // MathType 7 的 OLE 对象不包含此流；模板中残留的缓存（来自简单公式 "A"）
                // 会导致 MathType 在复杂公式上显示错误的预览图
                if (root.hasEntry("\u0002OlePres000")) {
                    root.getEntry("\u0002OlePres000").delete();
                }

                ByteArrayOutputStream baos = new ByteArrayOutputStream();
                fs.writeFilesystem(baos);
                log.debug("Packaged OLE using MathType template");
                return baos.toByteArray();
            }
        } catch (Exception e) {
            log.warn("Failed to package OLE by template, fallback to generated streams", e);
            return null;
        }
    }

    /**
     * 创建 \001Ole 流 — 标准 OLE2 嵌入对象头。
     *
     * <p>该流固定 20 字节，布局如下：
     * <pre>
     * 偏移  大小  含义
     * 0     4     版本号 + 标志位（0x02000001 = OLE 2.0，嵌入式对象）
     * 4     4     保留标志（0）
     * 8     12    保留字段（全零）
     * </pre>
     *
     * <p>此流告诉 OLE 宿主（Word）这是一个嵌入式对象（Embedded），而非链接对象（Linked）。</p>
     *
     * @return 20 字节的 \001Ole 流数据
     */
    private byte[] createOleStream() {
        ByteBuffer buf = ByteBuffer.allocate(20).order(ByteOrder.LITTLE_ENDIAN);
        buf.putInt(0x02000001); // 版本 2.0 + 标志位：嵌入式对象
        buf.putInt(0x00000000); // 保留标志
        buf.putInt(0x00000000); // 保留
        buf.putInt(0x00000000); // 保留
        buf.putInt(0x00000000); // 保留
        return buf.array();
    }

    /**
     * 创建 \001CompObj 流 — COM 对象标识信息。
     *
     * <p>CompObj 流是 OLE 规范中定义的对象标识流，Word 通过读取此流来确定
     * 嵌入对象的类型和关联的编辑程序。该流包含三个关键信息：</p>
     * <ul>
     *   <li><b>AnsiUserType</b> — 人类可读的类型名称 "MathType 7.0 Equation"，
     *       在 Word 的"对象属性"对话框中显示</li>
     *   <li><b>AnsiClipboardFormat</b> — 剪贴板格式标识（此处为 0，表示无注册格式）</li>
     *   <li><b>AnsiProgID</b> — 程序标识符 "Equation.DSMT4"，Word 通过此字符串
     *       在注册表 HKCR 中查找对应的 COM 类，从而激活 MathType 编辑器</li>
     * </ul>
     *
     * <p>流布局：
     * <pre>
     * [28 字节头部: reserved(4) + version(4) + 系统信息(20)]
     * [AnsiUserType: 4 字节长度前缀 + ASCII 字符串 + NULL 终止符]
     * [AnsiClipboardFormat: 4 字节（0x00000000 = 无）]
     * [AnsiProgID: 4 字节长度前缀 + ASCII 字符串 + NULL 终止符]
     * </pre>
     *
     * @return CompObj 流的字节数据
     * @throws IOException 写入内部缓冲区失败时抛出
     */
    private byte[] createCompObjStream() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream(256);
        ByteBuffer buf;

        // === 28 字节头部 ===
        buf = ByteBuffer.allocate(28).order(ByteOrder.LITTLE_ENDIAN);
        buf.putInt(0xFFFFFFFF); // 保留字段（固定值）
        buf.putInt(0x00000002); // CompObj 流版本号
        // 20 字节系统信息：字节序(2) + OS版本(2) + OS类型(2) + 未使用(10) + 保留(4)
        buf.putShort((short) 0xFFFE); // 字节序标记：小端序（Intel）
        buf.putShort((short) 0x000A); // OS 版本
        buf.putInt(0x00000002);       // OS 类型：Win32
        for (int i = 0; i < 8; i++) buf.put((byte) 0); // 剩余保留字节填零
        out.write(buf.array());

        // AnsiUserType：人类可读的对象类型名称（在 Word 对象属性中显示）
        writeAnsiString(out, "MathType 7.0 Equation");

        // AnsiClipboardFormat：剪贴板格式（0 = 无特定注册格式）
        buf = ByteBuffer.allocate(4).order(ByteOrder.LITTLE_ENDIAN);
        buf.putInt(0x00000000);
        out.write(buf.array());

        // AnsiProgID：程序标识符，Word 通过此 ProgID 查找并激活 MathType
        writeAnsiString(out, "Equation.DSMT4");

        return out.toByteArray();
    }

    /**
     * 创建 "Equation Native" 流 — MathType 公式的核心数据流。
     *
     * <p>该流是 MathType OLE 对象中最关键的部分，包含公式的完整二进制描述。
     * 格式为 28 字节的 EQNOLEFILEHDR 头部 + MTEF 数据负载。</p>
     *
     * <p>EQNOLEFILEHDR 结构定义（所有字段均为小端序）：
     * <pre>
     * struct EQNOLEFILEHDR {
     *   WORD  cbHdr;     // 头部大小 = 28 字节
     *   DWORD version;   // 版本号，高字 = 2，低字 = 0 → 0x00020000
     *   WORD  cf;        // 剪贴板格式（0 = 无）
     *   DWORD cbObject;  // MTEF 数据的字节长度
     *   DWORD reserved1; // 保留（0）
     *   DWORD reserved2; // 保留（0）
     *   DWORD reserved3; // 保留（0）
     *   DWORD reserved4; // 保留（0）
     * };
     * </pre>
     *
     * <p>MTEF（MathType Equation Format）是 MathType 的原生二进制公式格式，
     * 包含公式树的完整结构：行、分数、根号、上下标、矩阵等各类节点。</p>
     *
     * @param mtefData MTEF v5 格式的二进制公式数据
     * @return 完整的 "Equation Native" 流字节数据（28 字节头 + MTEF 负载）
     */
    private byte[] createEquationNativeStream(byte[] mtefData) {
        int hdrSize = 28;
        ByteBuffer buf = ByteBuffer.allocate(hdrSize + mtefData.length).order(ByteOrder.LITTLE_ENDIAN);
        buf.putShort((short) hdrSize);      // cbHdr：头部大小 = 28
        buf.putInt(0x00020000);             // version：高字=2, 低字=0（MTEF 版本标识）
        buf.putShort((short) 0);            // cf：剪贴板格式（0 = 无特定格式）
        buf.putInt(mtefData.length);        // cbObject：MTEF 数据的字节长度
        buf.putInt(0);                      // reserved1（保留）
        buf.putInt(0);                      // reserved2（保留）
        buf.putInt(0);                      // reserved3（保留）
        buf.putInt(0);                      // reserved4（保留）
        buf.put(mtefData);                  // MTEF 二进制公式数据负载
        return buf.array();
    }

    /**
     * 基于模板头部构建 "Equation Native" 流。
     *
     * <p>复用模板中原有 EQNOLEFILEHDR 头部的全部字段（版本号、保留字段等），
     * 仅修改 cbObject（偏移 8，4 字节 DWORD）为新 MTEF 数据的长度，
     * 并替换头部之后的 MTEF 数据负载。</p>
     *
     * <p>如果模板头部不可用或格式异常，则退化为从零构建。</p>
     *
     * @param templateEquationNative 模板中原有的 "Equation Native" 流字节（可为 null）
     * @param mtefData               新的 MTEF 二进制公式数据
     * @return 组装后的 "Equation Native" 流字节数据
     */
    private byte[] createEquationNativeStreamFromTemplate(byte[] templateEquationNative, byte[] mtefData) {
        // 模板数据不可用或不足以包含完整头部时，退化为从零构建
        if (templateEquationNative == null || templateEquationNative.length < 28) {
            return createEquationNativeStream(mtefData);
        }

        // 从模板头部前 2 字节读取 cbHdr（头部大小）
        ByteBuffer hdrLenBuf = ByteBuffer.wrap(templateEquationNative, 0, 2).order(ByteOrder.LITTLE_ENDIAN);
        int hdrSize = Short.toUnsignedInt(hdrLenBuf.getShort());

        // 头部大小不合理时退化
        if (hdrSize < 12 || templateEquationNative.length < hdrSize) {
            return createEquationNativeStream(mtefData);
        }

        // 复制模板头部，追加新的 MTEF 数据
        byte[] out = new byte[hdrSize + mtefData.length];
        System.arraycopy(templateEquationNative, 0, out, 0, hdrSize);

        // 更新 cbObject 字段（EQNOLEFILEHDR 偏移 8 处的 DWORD）为新 MTEF 数据长度
        ByteBuffer.wrap(out, 8, 4).order(ByteOrder.LITTLE_ENDIAN).putInt(mtefData.length);

        // 将新的 MTEF 数据写入头部之后
        System.arraycopy(mtefData, 0, out, hdrSize, mtefData.length);
        return out;
    }

    /**
     * 创建 \003ObjInfo 流 — 嵌入对象的显示信息。
     *
     * <p>该流固定 6 字节，所有字段为 0 表示这是一个标准的嵌入式 OLE 对象：
     * <ul>
     *   <li>标志位 = 0：非链接对象，无图标显示</li>
     *   <li>保留字段 = 0</li>
     * </ul>
     *
     * @return 6 字节的 \003ObjInfo 流数据
     */
    private byte[] createObjInfoStream() {
        ByteBuffer buf = ByteBuffer.allocate(6).order(ByteOrder.LITTLE_ENDIAN);
        buf.putShort((short) 0x0000); // 标志位：嵌入式，无链接，无图标
        buf.putShort((short) 0x0000); // 保留
        buf.putShort((short) 0x0000); // 保留
        return buf.array();
    }

    /**
     * 写入带 4 字节长度前缀的 ANSI 字符串。
     *
     * <p>OLE CompObj 流中的字符串采用统一的编码方式：
     * 先写入 4 字节小端序的长度值（包含 NULL 终止符），
     * 再写入 ASCII 字节内容，最后追加一个 0x00 终止符。</p>
     *
     * @param out 输出流
     * @param str 要写入的字符串
     * @throws IOException 写入失败时抛出
     */
    private void writeAnsiString(ByteArrayOutputStream out, String str) throws IOException {
        byte[] bytes = str.getBytes(StandardCharsets.US_ASCII);
        ByteBuffer lenBuf = ByteBuffer.allocate(4).order(ByteOrder.LITTLE_ENDIAN);
        lenBuf.putInt(bytes.length + 1); // +1 包含 NULL 终止符
        out.write(lenBuf.array());
        out.write(bytes);
        out.write(0x00); // NULL 终止符
    }
}
