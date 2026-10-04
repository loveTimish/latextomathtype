# 受控 Linux PDF HTTP 导出

该功能默认关闭。它使用和 Word 路由完全一致的 DOCX 导出器，随后调用现有几何基线 UNO 兼容转换。不会接收客户端 DOCX 路径或任意可执行命令，也不保存/重写源 DOCX。MathType OLE/MTEF 与 WMF 保留在源 DOCX；LibreOffice 本身会将预览栅格化到 PDF，不能把最终 PDF 宣称为全矢量。

## 启用

先按 `scripts/pdf/README.md` 安装已验证官方 LibreOffice/Java UNO 依赖。运维配置必须指向可信、不可由请求用户修改的代码和二进制。项目不会自动下载或运行未知依赖。

```sh
java \
  -Dpaperword.assets.root=/srv/paperword/assets \
  -Dpaperword.mathjax.node.command=/opt/node/bin/node \
  -jar target/paper-to-word-1.0.0.jar \
  --paperword.pdf.enabled=true \
  --paperword.pdf.helper=/srv/paperword/scripts/pdf/export_with_baseline.py \
  --paperword.pdf.uno-jar=/srv/paperword/pdf-tools/libreoffice-24.8.0.jar \
  --paperword.pdf.soffice-command=/usr/bin/soffice
```

图片策略见 `image-assets-security.md`。图片 JVM `-D` /环境变量配置对 HTTP 和直接 Java/CLI 导出一致，不应误写为仅 Spring YAML 配置。

- POST `/api/export/pdf`：请求 JSON 同 `/api/export/word`
- POST `/api/export/layout-pdf`：请求 JSON 同 `/api/export/layout-word`
- 默认 `?mode=compatible` 恢复 DOCX 中的几何基线；`?mode=native_layout` 返回相同输入的原生 LibreOffice 对照
- 返回 `application/pdf`，只有全部转换和校验成功后才提交成功响应
- GET `/api/export/pdf/status`：enabled、active、maxConcurrent、completed、failed，不暴露服务器目录

两路内部均生成 native/compatible 供核对。服务仅返回所选 PDF；审计 TSV、原始 DOCX 和子进程日志放在该任务的私有临时目录，任务结束清理。需要长期审计或双 PDF 文件时，可直接使用 CLI，它会保留 PDF、TSV 和源 SHA256；转换失败不发布任何 PDF，仅保留至多 64 KiB 日志。

## 资源/故障约束

| 配置（前缀 paperword.pdf.） | 默认 | 范围/含义 |
|---|---:|---|
| enabled | false | 显式启用 |
| max-concurrent | 1 | 1–8；满时立即 429，没有无界等待队列 |
| timeout-seconds | 120 | 1–600；覆盖编译/LO启动/转换；服务另加5秒清理余量，不是整个DOCX构建的总deadline |
| max-docx-mib | 32 | 1–256；生成后转换前检查 |
| max-pdf-mib | 64 | 1–256；读取响应前检查 |
| python-command | python3 | 单独可执行文件，不作shell解析 |
| java-command | java | Java21含compiler module |
| soffice-command | soffice | Linux LibreOffice |
| helper / uno-jar | 无 | 必需的可信本地路径 |
| cjk-font | Noto Serif CJK SC | 兼容输出用于缺失东亚字体的已安装替代字体；空值关闭显式回退 |

每任务独立 profile 和仅 loopback UNO；禁用宏和链接更新；输入包外部关系被拒绝。服务用 /usr/bin/setsid 为每任务建立单一私有进程组，编译器/LO/UNO 都留在该组；Java 用 /bin/kill 对整组做最终回收，即使 helper 卡住或 LO launcher 已退出也覆盖孤儿进程。CLI 单独运行时各子进程使用私有会话组自行清理。Java 服务设外层deadline并临时清除中断标志完成有界清理后恢复。日志有64KiB上限，源DOCX校验SHA不变。无公式普通文档正常导出，不凭空增加对象。

错误 JSON 为 `code` 和 `message`：503 PDF_DISABLED/PDF_UNAVAILABLE/PDF_CONFIGURATION；429 PDF_BUSY；504 PDF_TIMEOUT；413 PDF_INPUT_TOO_LARGE/PDF_OUTPUT_TOO_LARGE；422 PDF_CONVERSION_FAILED。图片错误统一由 Word/PDF 返回400或413，未配置图片根为503。无成功PDF可用于部分/失败结果。生产仍需网关请求体限额、认证/租户隔离、总请求deadline和OS容器内存/CPU配额；本改动不把演示服务变成可直接公开部署的完整平台。

## 自动回归

```sh
python scripts/pdf/test_export_with_baseline.py -v
# 再运行 Maven: LinuxPdfExportServiceTest,PdfExportControllerTest
```

前者用独立合成进程覆盖成功/不覆盖原文件/半成品失败/超时/SIGTERM/缺依赖；后者覆盖服务slots、取消、错误状态及输出检查。它们不能替代真实LibreOffice视觉验收。交付报告记录通过实际HTTP入口完成的固定样本、PNG逐页对照及原生Word/MathType未测边界。

## 明确的 CJK 字体回退

真实仓库 `exam-template.json` 揭示环境问题：机器没有宋体时，LibreOffice 在部分节标题错误选择 Noto Sans Mongolian，导致正文XML存在但PDF缺字符。该问题在有界缓存与资产策略修改前的阶段快照同样存在。

兼容输出现在在导出原生对照后，对包含CJK文字、且原请求的 `CharFontNameAsian` 不在LibreOffice可用字体清单中的Text portions，显式设置已安装的 `paperword.pdf.cjk-font`。默认 Noto Serif CJK SC，保持已安装的用户字体不变，不改正文、非CJK字体、字体大小或OLE。映射与次数写入 `CJK_FONT_FALLBACK` 审计日志。需要回退却找不到指定替代字体时明确503，不返回缺字的兼容PDF。CLI对应 `--cjk-font`。

此类真实字体替换可能改变自然行高、换行和页数。仓库六题模板从不完整的原生1页变为完整兼容2页；没有硬编码行高或缩小字号压回1页。原生 `mode=native_layout` 始终保存未经字体或基线修复的LO对照，仍可能缺字/偏高，不作为推荐成品。原始DOCX宋体声明与源SHA保持不变。

服务两种mode都会完成原生+兼容双输出及共同校验后再返回所选文件；因此缺少必须的CJK替代字体时，native_layout也会明确503。若只需旧行为的诊断，CLI可显式传空`--cjk-font`，并承担缺字风险。

字体名称策略是有限且失败关闭的，并不重写LibreOffice完整字体解析器：支持逗号/分号family列表、feature后缀、清单中的规范化名称和一组有官方依据的常见CJK别名。任一实际可验证的已安装候选均保留；只有已知family/alias确实缺失才替代，其他未能安全判断的自定义/本地化/PostScript名称明确503，避免覆盖未知的已安装别名或返回缺字PDF。页眉/页脚只处理在用且已启用的文本区域，PAGE/NUMPAGES字段的指令和对象不变，合法分页可更新显示页码。
