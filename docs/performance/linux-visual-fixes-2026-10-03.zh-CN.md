# 第二轮：行内基线与混合长除法实际修复

第二阶段继续在同一 Linux 工作区实施，保留第一阶段的并发隔离、single-flight、超时回收修复。本文是该阶段的历史测量，资产安全与缓存限额的后续修复见对应配置文档。

## 已完成

1. **混合长除法不再丢前后项**。预览从完整 MathIR 表达式生成 MathML，保留兄弟节点、多处长除法、分式、根号、上下标、定界符及支持的样式；未知结构明确报错。修复 indexed-root 顺序、dbinom/tbinom 样式和大幅 raisebox 的边界框。纯数字长除法的既有序列化保持一致，MTEF 写入逻辑未改。
2. **DOCX 的 baseline 改为真实几何**。移除分式取 75% 高度等类型猜值，把 MathJax baseline 经与字形完全相同的 Batik→WMF 仿射变换映射到目标框，包含等比缩放、居中和安全留白。未知 baseline 明确拒绝，避免把 -1 误算成有效下沉量。
3. **XML 与字号修复**。OLE 正确放在直接 `w:r/w:object` 中，消除嵌套 run；保留 POI run 的有效引用，失败不再被吞掉。Layout API 中公式与 baseline 随明确的 11/16pt 正文字号同比例变化，没有通过缩小公式掩盖错位。
4. **提供显式 Linux PDF 兼容导出**。源码和实测确认 LibreOffice 导入 OLE 时忽略 `w:position`、强制对象底边贴正文基线。脚本通过官方 UNO API 恢复度量后的 AS_CHARACTER 定位，再正常导出 PDF；同时保存未经调整的原生 PDF。源 DOCX 不重写，全部 OLE 保留，未编辑 PDF 像素或替换成截图。

UNO 使用的几何关系是：

`position = -objectHeight - topMargin + descent`

全部统一到 1/100mm，bottomMargin 保留为行间留白。按正文及表格真实流顺序映射所有对象，并校验总数、顺序、尺寸、inline anchor；设置后断言宽高不变。只支持生成器已覆盖的水平、非旋转、内嵌 MathType 对象。External 关系、链接对象、未知 baseline 或无法完整映射的输入在加载/导出前失败。宏、链接更新及 XML 外部实体均禁用。

## 真实视觉验收

独立验收使用同一套固定合成输入，经过真实 LibreOffice 转换、Poppler 页面渲染，并逐页检查：

- 复杂式/有效图片 1 页
- 11pt、16pt、中西文行内、display、16pt 表格 1 页
- 原表格样本 1 页
- 原 40 公式样本 6 页

原生与兼容路径各 9 页。兼容路径均无裁切、漏式、碰边、相邻行重叠或新增空白页。混合长除法左右的 `x+`、`+1` 在默认原生 PDF 中也已完整出现，说明这是源预览修复。

首行 `x²+1` 末尾数字底边相对邻近正文 baseline：

| 相同最终 DOCX 的转换方式 | 可见底边误差 |
|---|---:|
| 原生 LibreOffice | 约 -2.958pt，偏高 |
| 显式 baseline 兼容路径 | 约 +0.042pt |

300dpi 采样精度约 0.24pt。该公式在两路 PDF 中的图像像素和显示尺寸相同，因此改善来自定位，而非缩小字号或重画公式。40 式仍为 6 页，题号范围因基线恢复自然重排为 1–7 / 8–15 / 16–22 / 23–30 / 31–37 / 38–40。

复杂 DOCX 与旧样本相比：11 个 OLE embedding 字节全部不变；只改变第 8 个混合长除法 WMF，其余 10 个 WMF 与有效 PNG 不变。最终重导出另与已验收候选逐 ZIP part 核对，排除创建时间元数据差异，确保交付代码对应已检查内容。

## 测试与审查

- 最终 Java 选定套件：**486 项，479 通过、0 失败、0 错误、7 条件跳过**
- 范围仍为 `src/test/java` 下除依赖外部语料/Office 的 `tools/` 产物生成器外的 `*Test.java`，不宣称全库默认 `mvn test` 无条件通过
- MathJax smoke 与 npm 几何回归全部通过；可执行 JAR 打包成功
- 新回归包含实际 SVG baseline、目标框 letterboxing、安全边距、unknown sentinel、直接 OLE run、11/16pt 比例、完整长除表达式及 MTEF 非变更
- 大幅 raisebox 上下 60pt 的真实 MathJax 回归确认所有字形位于 viewBox 内；binomial display/text 样式区分保留
- 独立代码复审提出的 unknown baseline、raisebox 裁切、binomial 样式、外部关系检查均已闭合；未发现新的阻塞问题
- 前轮共享服务并发、同公式冷 miss 合并、失败重试、超时后恢复和等待者中断回归继续通过

## 必须保留的边界

- **原生 LibreOffice 导入器问题仍存在**。DOCX 几何修复与额外 PDF 兼容转换是两项不同交付；不能说单靠 DOCX 就让所有转换器都正确
- **PDF 不是全矢量公式输出**。此 LibreOffice 版本在原生和兼容路径均把 OLE 预览内部栅格化约 300dpi；DOCX 中仍是可编辑 OLE/MTEF 加纯矢量 WMF。UNO没有新增位图替换。某个竖式的 PDF 栅格采样有约 0.1pt 量化差异，未见裁切，不能泛称所有 PDF 像素完全不变
- Windows Word/MathType 原生双击编辑、保存重开与视觉验收尚未执行
- 第一轮发现的任意本地配图路径读取、内存/磁盘缓存无上限仍未在本轮解决；对不可信请求开放前应处理
- MathIRConverter 历史上丢弃显式间距命令（如 `\quad`）的限制仍在；本轮不扩展原始解析器语义

## 交付与复现

- 当前代码包含第一阶段的性能/并发修复；历史上游基线为 `bf56c59bc823a8a1bab518b4bc3a3d13d1a113d5`
- 复测应同时保留原生与兼容转换输出，并核对 OLE 数量、几何基线和实际页面
- 脚本及准确命令：`scripts/pdf/README.md`。依赖官方 `org.libreoffice:libreoffice:24.8.0` Java UNO jar；不启用宏、不降低安全设置

LibreOffice 实现依据：
- [OLE 导入](https://github.com/LibreOffice/core/blob/master/sw/source/writerfilter/dmapper/DomainMapper_Impl.cxx)
- [VML 替换预览导入](https://github.com/LibreOffice/core/blob/master/oox/source/vml/vmlshape.cxx)
- [行内对象定位](https://github.com/LibreOffice/core/blob/master/sw/source/core/objectpositioning/ascharanchoredobjectposition.cxx)
- [实际 frame 定位](https://github.com/LibreOffice/core/blob/master/sw/source/core/layout/flyincnt.cxx)
- [UNO 属性定义](https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1text_1_1BaseFrameProperties.html)
