# XSC LaTeX Assets

The XSC source set for this checkout is:

```text
E:\新加卷\新建文件夹\xsc资料
```

Convert the source DOCX files into fixed LaTeX assets before rebuilding Word
documents with `latextomathtype`. The local release package is:

```text
H:\下载\docx2tex-1.10-release (2).zip
```

It is unpacked to:

```text
J:\latextomathtype\analysis\tools\docx2tex-1.10-release\docx2tex
```

Run:

```powershell
.\scripts\convert_xsc_docx_to_latex_assets.ps1 `
  -SourceDir 'E:\新加卷\新建文件夹\xsc资料' `
  -OutRoot 'J:\latextomathtype\analysis\xsc-latex-fixed' `
  -Recurse
```

The output is a fixed numbered asset set:

```text
J:\latextomathtype\analysis\xsc-latex-fixed\manifest.json
J:\latextomathtype\analysis\xsc-latex-fixed\failures.json
J:\latextomathtype\analysis\xsc-latex-fixed\<index>\<index>.tex
```

Use the manifest to map every stable index back to the source document name.
