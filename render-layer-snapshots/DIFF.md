# 本地与远端渲染层差异速览

## 远端 `main`

- `MathTypeEmbedder` 统一走 OLE 预览链，没有单独把长除法分流成图片。
- `LaTeXImageRenderer` 以普通公式预览为主，核心是 `MathJax HTTP -> dvisvgm -> JLaTeXMath` 的通用回退链。
- 远端 `LaTeXImageRenderer` 没有你现在这套长除法专用入口，也没有 `xlop`、`VerticalLayoutCompiler`、结构化竖式绘制这条支线。

## 本地当前

- `MathTypeEmbedder` 多了长除法识别，命中后直接走 `insertLongDivisionPicture(...)`。
- `LaTeXImageRenderer` 多了 AST/竖式编译、结构化预览、长除法专用图片渲染、`xlop` 官方链尝试、本地兜底绘制等一整套扩展。
- 本地还多了 MiKTeX 命令探测、`renderForOlePreview(LaTeXNode, String)` 重载、长除法操作数正规化等逻辑。

## 当前结论

- 如果你的目标是“先看远端原始渲染层是什么样”，那远端的关键特点就是：普通公式统一渲染，没有本地现在这层长除法分流。
- 如果后面要做“回退到远端基线，再单独补长除法”，现在这个快照目录已经足够作为第一轮对照入口。
