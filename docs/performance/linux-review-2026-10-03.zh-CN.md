# Linux review、性能基线与优化报告

日期：2026-10-03 UTC

## 结论

本次使用合成输入完成 Linux 源码审查、真实 Java/MathJax 渲染、DOCX 导出、并发复现、修复和回归。

这是第一阶段的历史测量报告。文中当时未修复的资产读取、混合长除法和缓存容量问题已由后续改动处理；当前配置及限制请参阅 `image-assets-security.md`、`render-cache.md` 和 `linux-pdf-service.md`。

最重要的结果是**修复并发文档的 MTEF 状态串扰**，其次是合并相同冷公式的重复渲染。8 个并发请求同一未缓存公式，15 批实测延迟中位数从 **153.45 ms 降到 61.32 ms，约 2.50 倍**；实际 MathJax 请求从每批 8 次减为 1 次。没有通过禁用 WMF、降低公式精度、改变几何拟合或替换为截图来提速。

**不能据此宣称所有场景加速，也还不能作为对不可信请求开放的生产发布验收。** 任意本地文件配图路径和混合长除法预览仍有明确缺陷，详见下文。

## 版本与环境

- 仓库：`https://github.com/loveTimish/latextomathtype`
- 固定基线：`bf56c59bc823a8a1bab518b4bc3a3d13d1a113d5`
- 本地分支：`perf/linux-export-review`
- Debian GNU/Linux 13，Linux 6.18.44，x86_64
- CPU：Intel Xeon Platinum 8573C，容器可见 9 个逻辑 CPU；共享云环境，非独占物理机
- 可见内存约 9.73 GiB，无 swap；性能 JVM 固定 `-Xms256m -Xmx1024m`
- OpenJDK 21.0.12.1；Maven 3.9.12；Node **24.9.0**；MathJax 3.2.2；Saxon-JS 2.7.0
- 最初系统 Node 为 24.19.0；所有正式渲染/性能测试均使用工作区安装的 24.9.0，未绕过版本或 bundle-hash 校验
- 基础性能关闭磁盘缓存、保留原有 JVM 内存缓存；另单独验证了跨 JVM 磁盘缓存
- RSS 每 25 ms 采样，可能漏过更短峰值；heap-used 快照不是稳定存活堆或泄漏测量

## 已修复且有实证的问题

### 1. P1：服务共享可变导出器，发生 MTEF 状态串扰

基线位置：`PaperExportService.java:20`、`LayoutExportService.java:25`；`DocxBuilder.java:198` 重置共享公式编号；`MathTypeEmbedder.java:44` 复用有状态 `MtefWriter`。

同一个服务实例，8 线程、8 批，共 64 份文档，每份 12 个公式：基线 **63/64** 份与同请求串行参考的 MTEF 不一致，修复后 **0/64**。比较的是从 OLE `Equation Native` 中提取的 MTEF，不包含 ZIP 时间戳、OLE 外壳或对象 ID。

保留的证据文档中，12 个 OLE 都可解析，但只有 4 个通过规范化 MTEF 精确结构比较；第 2 式的第 25 个规范记录，串行参考为 `FULL`，并发输出为 `LINE`。这是写入状态串扰；不将每个二进制差异都解释成数学字符被替换。

修复：每次服务导出创建自己的 builder，文档编号、布局标志、MTEF writer 都归本次导出所有；渲染缓存仍可共享。两种 HTTP 导出入口都经过对应服务。直接使用 builder 的库调用方仍应遵循“一个并发导出一个 builder”的线程约束。

### 2. P1：worker 超时清理锁住，响应超时失效

基线位置：`LaTeXImageRenderer.java:1022–1059`。异步 `readLine()` 持有 reader 锁，超时处理却先 `close()` reader，再杀子进程，导致清理本身等待。

用无响应的合成 Node worker，将响应超时设为 1 秒：基线超过外层 8 秒上限，进程被测试器终止（exit 124）；修复后约 **1.10 秒**返回预期异常。回归另验证超时后下一次真实渲染成功。

修复：先解除静态引用并杀 worker，再关闭流；reader/stderr 任务捕获本轮实例，避免读取替换后的全局对象；处理中断并恢复中断标志。

### 3. 性能：同一冷公式并发 miss 重复渲染

基线位置：`LaTeXImageRenderer.java:196–241`。多个调用均在首个结果写入缓存前 miss，分别执行同一公式的 MathJax/Batik/WMF 工作。

修复：按完整现有缓存键合并正在执行的 OLE 预览请求。不同键不共用这个等待结果；失败不进入缓存；在 finally 删除正在执行的条目。等待者取消不会取消其他调用需要的共享工作，等待者异常与 owner 异常保持一致。

### 4. P2：默认 locale 改变 VML 尺寸小数点

基线位置：`MathTypeEmbedder.java:178–179`。`Locale.GERMANY` 下实测生成 `width:12,345pt` 类尺寸。改为 `Locale.ROOT`，回归验证只生成点分隔的 VML 尺寸。

### 5. Linux 可复现构建问题

- Maven shell 脚本原来是 100644，README 的直接执行命令会失败；仅改为 100755，没有更改正文
- 官方参考 HTML 在 Linux 的 LF 检出与 corpus 中固定的 CRLF 原始 SHA 不一致。确认只有 8 个换行字节差异、正文相同后，只为该一个参考 HTML 指定 `eol=crlf`。没有改 corpus hash 或任何用户内容

## 性能数据

同一机器、同一输入、相同堆配置、相同 Node/MathJax bundle。前后版本顺序交替，CPU 密集测试没有并行运行。下表为中位数，除特别注明外均为 3 个独立 JVM。

| 场景 | 基线 | 优化后 | 说明 |
|---|---:|---:|---|
| 8 并发同一冷复杂公式 | 153.45 ms | 61.32 ms | 每 JVM 5 批，共 15 批；8 次实际渲染降到 1 次 |
| 上述批次等效请求吞吐 | 52.1 请求/s | 130.5 请求/s | 8 / 中位批次耗时；不是全服务饱和吞吐 |
| 上述批次样本 P95 | 369.58 ms | 81.30 ms | 仅 15 个样本，不能当生产 SLO |
| 第一个简单公式，冷 JVM/worker | 2503.05 ms | 2351.88 ms | 包含类加载、字体与 worker 初始化 |
| 同公式内存命中 1000 次 | 208.84 ms | 160.61 ms | 不代表全 DOCX 导出速度 |
| 40 个不同复杂公式，初期 JVM | 2829.92 ms | 3352.51 ms | 此轮中位数变慢，保留结果、不隐藏 |
| 充分预热后 60 个不同复杂公式 | 3352.44 ms | 3358.86 ms | **5 个独立 JVM**，每个先预热 20 个不同公式；差约 +0.19% |
| 冷渲染缓存，40 题/40 不同公式完整 DOCX | 5849.20 ms | 7137.34 ms | 首轮中位数约 +22%；初始化/共享环境波动明显 |
| 同一 40 题文档第 2 次导出 | 305.78 ms | 305.37 ms | JVM 渲染缓存已热 |
| 同一 40 题文档第 3 次导出 | 298.40 ms | 243.09 ms | JVM 渲染缓存已热 |
| 冷缓存 100 题/100 不同公式 DOCX | 10469.84 ms | 9869.50 ms | 每版本仅 1 JVM，只作为规模 smoke |
| performance 全场景 JVM 峰值 RSS | 375.23 MiB | 366.30 MiB | 3 JVM 峰值的中位数 |
| performance 全场景 JVM+Node 峰值 RSS | 534.21 MiB | 524.28 MiB | 采样近似值，非容量保证 |

结论：可确认的性能收益集中于**同公式并发冷 miss**。不同公式的稳态吞吐基本持平；小样本冷启动/冷整卷数据波动且并非全都改善，应在目标 Linux 主机、真实题卷分布和并发量下重新做发布容量测试。

HTTP 主进程冷启动到 health 的 3 次中位数为 2587.65→3043.07 ms；首次两公式带 PNG 的 HTTP 导出 3599.15→2988.64 ms；紧接着的热导出 73.75→71.61 ms。主进程启动测试使用同一 Spring Boot 主类 classpath；另对打包后的 executable JAR 做实际启动 smoke。没有将 classpath 数字伪称为 JAR 启动基准。

跨 JVM 磁盘缓存额外实测（各一组，非统计基准）：基线 40 式首次 6160.01 ms、重启后 1059.29 ms；优化后 6421.74 ms、重启后 1284.61 ms。两个版本重启后的 worker 请求数均为 **0**，说明现有持久化缓存命中仍有效。

## 正确性与回归

- 最终选定 Java 套件：**467 项，460 通过、0 失败、0 错误、7 条件跳过**
- 范围：`src/test/java` 下除 `tools/` 语料/Office 产物生成器以外的 `*Test.java`，包含 controller、两种 service、parser、MathML、MTEF、OLE、Batik/WMF、DOCX、Linux 回归
- 原始相同范围测试有固定快照换行哈希失败；修复后通过。最初全量运行还暴露一个未入库的参考 JSON 缺失，并在中断时未完成，因此**不报告全库 mvn test 全通过**
- Mockito 动态自附着在此容器不可用；测试 JVM 改为启动时加载已下载的 Byte Buddy agent，没有修改 OS 安全设置
- `npm run mathjax:smoke` 与 `npm test` 全部通过
- `mvn -DskipTests package` 成功；打包 JAR 的 Linux health smoke 成功
- 前后 30 份成对 DOCX，合计 **1380 个公式**：规范 MTEF/原始 MTEF 精确比较、预览媒体逐字节比较、document.xml 逐字节比较；所有配对一致
- 各性能输出检查 OLE 公式数、Equation Native/DSMT 标识、MTEF 可解析性、WMF 预览数和纯矢量记录；未使用 PNG 替换公式
- 6 次真实 HTTP 启动、共 12 次导出，全部包含 2 OLE、2 WMF 和 1 张与输入逐字节一致的合成 PNG
- 新回归覆盖两个服务并发、相同冷 miss 合并、失败重试、超时后恢复、等待者中断和德语 locale
- 独立只读复审未发现当前生产补丁引入的新缺陷；复审提出的 waiter 中断与测试退出清理问题已修正并复测

## 未解决的确定缺陷与边界

### P1：配图路径能读出服务账号可读的任意本地文件

基线 `DocxBuilder.java:644–715`、`LayoutDocxBuilder.java:179–204` 接受调用方路径，未限定资源根目录。普通导出中，即使 ImageIO 无法识别为图片，也会继续把原始流写入 DOCX。

本次只创建不含隐私的 `SYNTHETIC_TEST_CANARY_NO_PRIVATE_DATA` 文本，并把它作为 images 路径传入；导出的 `word/media/` 中确实出现完整文本。没有读取系统秘密或用户文件。**对不可信调用方开放前必须修复**：明确允许的资产根目录、realpath/符号链接策略、常规文件与实际图像校验、尺寸/字节上限；需要保留现有 Windows 路径映射与合法资产流程。本补丁没有擅自更改这项输入契约。

### P1：同一个公式内混合 longdivision 时，预览丢掉其他节点

基线 `LaTeXImageRenderer.java:976–989` 只从解析树取第一个 LONG_DIVISION 写 MathML。实测“单独长除法”和“x+长除法+1”的 SVG **逐字节完全相同**。MTEF 写入仍处理完整公式，造成公式正文与可见预览不一致。未在本轮性能补丁中扩展长除法组合排版；发布前应支持完整组合或明确拒绝此类输入，不能静默丢项。

### P2：缓存资源仍无上限

`LaTeXImageRenderer` 的静态 PREVIEW_CACHE、PNG_CACHE 以及磁盘缓存没有淘汰策略。此为源码可确认的资源管理限制；本次未进行数小时/海量独特公式的 OOM 压测，也没有声称测出长期泄漏。建议后续增加按字节权重的上限、磁盘配额/过期和可观测命中率。

响应超时修复没有增加整个排队/大输入 stdin 写入的总 deadline；未做超大请求、恶意资源耗尽、目标生产部署、Docker 镜像或离线 sidecar 完整发行打包验收。

**Linux 无法完成 Windows Word+MathType 双击编辑、保存再打开与视觉 GUI 验收。** OLE/MTEF 和预览结构通过，不等于原生 GUI 兼容性验收完成。

## 交付与复现

- `scripts/linux-review/`：Java 合成负载、Python 重复 JVM/RSS 采样、HTTP smoke 和复现说明
- 性能数据为上述合成负载的历史测量；目标部署应重新运行脚本建立容量基线
- 样例仅含合成公式和合成图像
- 应用方法及固定参考 HTML 的检出步骤见 `scripts/linux-review/README.md`
