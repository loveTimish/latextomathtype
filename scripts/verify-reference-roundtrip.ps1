param(
    [string]$ReferenceDocx = "rebuild-assets\external\fraction-split-reference.docx",
    [string]$Docx2TexRoot = "J:\docx2tex",
    [string]$OutDir = "target\reference-roundtrip",
    [string]$FormulaDatasetRoot = "E:\新加卷\新建文件夹\xsc资料",
    [int]$FormulaDatasetMaxDocs = 40,
    [int]$MinimumOleCount = 400,
    [double]$MinimumFormulaHeightRatio = 0.85,
    [double]$MaximumFormulaWidthOverflowPt = 5.0,
    [switch]$SkipFormulaDataset,
    [switch]$SkipVisual,
    [switch]$SemanticOnly,
    [switch]$CheckRenderedFormulaInk,
    [switch]$UseReferenceLayoutShell
)

$ErrorActionPreference = "Stop"

function Resolve-ProjectPath([string]$PathValue) {
    if ([IO.Path]::IsPathRooted($PathValue)) {
        return $PathValue
    }
    return Join-Path (Get-Location) $PathValue
}

function Read-ZipTextEntry([string]$ZipPath, [string]$EntryName) {
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [IO.Compression.ZipFile]::OpenRead($ZipPath)
    try {
        $entry = $zip.Entries | Where-Object { $_.FullName -eq $EntryName } | Select-Object -First 1
        if (-not $entry) {
            return $null
        }
        $reader = [IO.StreamReader]::new($entry.Open(), [Text.Encoding]::UTF8)
        try {
            return $reader.ReadToEnd()
        } finally {
            $reader.Dispose()
        }
    } finally {
        $zip.Dispose()
    }
}

function Get-DocxMetrics([string]$DocxPath) {
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [IO.Compression.ZipFile]::OpenRead($DocxPath)
    try {
        $names = @($zip.Entries | ForEach-Object { $_.FullName })
        $documentXml = Read-ZipTextEntry $DocxPath "word/document.xml"
        $stylesXml = Read-ZipTextEntry $DocxPath "word/styles.xml"
        $settingsXml = Read-ZipTextEntry $DocxPath "word/settings.xml"

        $media = @($names | Where-Object { $_.StartsWith("word/media/") })
        $embeddings = @($names | Where-Object { $_.StartsWith("word/embeddings/") })
        $shapeMatches = [regex]::Matches($documentXml, '<v:shape[^>]*style="([^"]+)"')
        $widths = New-Object System.Collections.Generic.List[double]
        $heights = New-Object System.Collections.Generic.List[double]
        foreach ($match in $shapeMatches) {
            $style = $match.Groups[1].Value
            $widthMatch = [regex]::Match($style, "width:([0-9.]+)(pt|in|cm|mm)")
            $heightMatch = [regex]::Match($style, "height:([0-9.]+)(pt|in|cm|mm)")
            if ($widthMatch.Success -and $heightMatch.Success) {
                $widths.Add((Convert-StyleUnitToPt ([double]$widthMatch.Groups[1].Value) $widthMatch.Groups[2].Value))
                $heights.Add((Convert-StyleUnitToPt ([double]$heightMatch.Groups[1].Value) $heightMatch.Groups[2].Value))
            }
        }

        $mediaExt = @{}
        foreach ($item in $media) {
            $ext = [IO.Path]::GetExtension($item).TrimStart(".").ToLowerInvariant()
            if (-not $ext) {
                $ext = "<none>"
            }
            if (-not $mediaExt.ContainsKey($ext)) {
                $mediaExt[$ext] = 0
            }
            $mediaExt[$ext]++
        }

        return [PSCustomObject]@{
            Path = $DocxPath
            Bytes = (Get-Item -LiteralPath $DocxPath).Length
            OleObjects = [regex]::Matches($documentXml, "<o:OLEObject\b").Count
            WordObjects = [regex]::Matches($documentXml, "<w:object\b").Count
            Embeddings = $embeddings.Count
            Media = $media.Count
            MediaExt = ($mediaExt.GetEnumerator() | Sort-Object Name | ForEach-Object { "$($_.Name)=$($_.Value)" }) -join ","
            Drawings = [regex]::Matches($documentXml, "<w:drawing\b").Count
            Paragraphs = [regex]::Matches($documentXml, "<w:p\b").Count
            Sections = [regex]::Matches($documentXml, "<w:sectPr\b").Count
            ShapeCount = $widths.Count
            ShapeWidthAvgPt = if ($widths.Count) { [math]::Round(($widths | Measure-Object -Average).Average, 2) } else { 0 }
            ShapeHeightAvgPt = if ($heights.Count) { [math]::Round(($heights | Measure-Object -Average).Average, 2) } else { 0 }
            ShapeWidthMaxPt = if ($widths.Count) { [math]::Round(($widths | Measure-Object -Maximum).Maximum, 2) } else { 0 }
            ShapeHeightMaxPt = if ($heights.Count) { [math]::Round(($heights | Measure-Object -Maximum).Maximum, 2) } else { 0 }
            HasStyles = [bool]$stylesXml
            HasSettings = [bool]$settingsXml
        }
    } finally {
        $zip.Dispose()
    }
}

function Convert-StyleUnitToPt([double]$Value, [string]$Unit) {
    switch ($Unit) {
        "in" { return $Value * 72.0 }
        "cm" { return $Value * 72.0 / 2.54 }
        "mm" { return $Value * 72.0 / 25.4 }
        default { return $Value }
    }
}

function Export-DocxPdf([string]$DocxPath, [string]$PdfPath) {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    try {
        $doc = $word.Documents.Open((Resolve-Path -LiteralPath $DocxPath).Path, $false, $true)
        try {
            $doc.ExportAsFixedFormat((Resolve-ProjectPath $PdfPath), 17)
        } finally {
            $doc.Close($false)
        }
    } finally {
        $word.Quit()
    }
}

function Export-PdfPages([string]$PdfPath, [string]$PageDir) {
    $pdftoppm = (Get-Command pdftoppm -ErrorAction SilentlyContinue)
    if (-not $pdftoppm) {
        throw "Missing pdftoppm; install MiKTeX/Poppler or run with -SkipVisual"
    }
    New-Item -ItemType Directory -Force -Path $PageDir | Out-Null
    Get-ChildItem -LiteralPath $PageDir -Filter "*.png" -ErrorAction SilentlyContinue |
        Remove-Item -Force -ErrorAction SilentlyContinue
    & $pdftoppm.Source -png -r 144 $PdfPath (Join-Path $PageDir "page")
    if ($LASTEXITCODE -ne 0) {
        throw "pdftoppm failed for $PdfPath with exit code $LASTEXITCODE"
    }
}

function Compare-VisualPages([string]$ReferenceDir, [string]$GeneratedDir) {
    $script = @'
import json
from pathlib import Path
from PIL import Image, ImageChops, ImageStat

ref_dir = Path(r"__REF_DIR__")
gen_dir = Path(r"__GEN_DIR__")
rows = []
for ref in sorted(ref_dir.glob("page-*.png")):
    gen = gen_dir / ref.name
    if not gen.exists():
        rows.append({"page": ref.name, "sameSize": False, "referenceSize": None, "meanDiff": None})
        continue
    with Image.open(ref).convert("RGB") as a, Image.open(gen).convert("RGB") as b:
        same_size = a.size == b.size
        if not same_size:
            b = b.resize(a.size)
        diff = ImageChops.difference(a, b)
        stat = ImageStat.Stat(diff)
        rows.append({
            "page": ref.name,
            "sameSize": same_size,
            "referenceSize": list(a.size),
            "meanDiff": round(sum(stat.mean) / len(stat.mean), 2),
            "diffBox": list(diff.getbbox()) if diff.getbbox() else None,
        })
print(json.dumps(rows, ensure_ascii=False, indent=2))
'@
    $script = $script.Replace("__REF_DIR__", (Resolve-ProjectPath $ReferenceDir).Replace("\", "\\"))
    $script = $script.Replace("__GEN_DIR__", (Resolve-ProjectPath $GeneratedDir).Replace("\", "\\"))
    $json = $script | python -
    if ($LASTEXITCODE -ne 0) {
        throw "visual page comparison failed with exit code $LASTEXITCODE"
    }
    return $json | ConvertFrom-Json
}

$referencePath = Resolve-Path -LiteralPath (Resolve-ProjectPath $ReferenceDocx)
$docx2TexPath = Resolve-Path -LiteralPath $Docx2TexRoot
$d2t = Join-Path $docx2TexPath "d2t.bat"
$conf = Join-Path $docx2TexPath "conf\conf.xml"
$mvn = Resolve-Path -LiteralPath ".mvn\apache-maven-3.9.12\bin\mvn.cmd"
$docx2TexOut = Join-Path $OutDir "docx2tex"
$generatedDocx = Join-Path $OutDir "fraction-split-reference-regenerated.docx"
$shellGeneratedDocx = Join-Path $OutDir "fraction-split-reference-regenerated-shell.docx"
$generatedJson = Join-Path $OutDir "fraction-split-reference.request.json"
$referenceFormulaBoxes = Join-Path $OutDir "formula-boxes-reference.json"
$uncalibratedFormulaBoxes = Join-Path $OutDir "formula-boxes-generated-uncalibrated.json"
$calibratedFormulaBoxes = Join-Path $OutDir "formula-boxes-generated-calibrated.json"
$formulaDatasetJson = Join-Path $OutDir "formula-box-dataset.json"
$formulaDatasetText = Join-Path $OutDir "formula-box-dataset.txt"
$formulaFitJson = Join-Path $OutDir "formula-box-fit-model.json"
$formulaFitText = Join-Path $OutDir "formula-box-fit-model.txt"
$formulaVectorSyncJson = Join-Path $OutDir "formula-vector-preview-sync.json"
$referenceLayoutShellJson = Join-Path $OutDir "reference-layout-shell-sync.json"
$wordFormulaLayoutJson = Join-Path $OutDir "word-formula-layout-comparison.json"
$wordFormulaLayoutText = Join-Path $OutDir "word-formula-layout-comparison.txt"
$renderedFormulaInkJson = Join-Path $OutDir "rendered-formula-ink-comparison.json"
$renderedFormulaInkText = Join-Path $OutDir "rendered-formula-ink-comparison.txt"
$referencePdf = Join-Path $OutDir "fraction-split-reference-current.pdf"
$generatedPdf = Join-Path $OutDir "fraction-split-reference-regenerated-current.pdf"
$referencePages = Join-Path $OutDir "visual-reference-current"
$generatedPages = Join-Path $OutDir "visual-generated-current"

if (-not (Test-Path -LiteralPath $d2t)) {
    throw "Missing docx2tex runner: $d2t"
}
if (-not (Test-Path -LiteralPath $conf)) {
    throw "Missing docx2tex config: $conf"
}

New-Item -ItemType Directory -Force -Path $docx2TexOut | Out-Null
& $d2t $referencePath $conf $docx2TexOut
if ($LASTEXITCODE -ne 0) {
    throw "docx2tex failed for reference document with exit code $LASTEXITCODE"
}

python rebuild\generate_fraction_reference_request.py
if ($LASTEXITCODE -ne 0) {
    throw "PaperExportRequest generation failed with exit code $LASTEXITCODE"
}
if (-not (Test-Path -LiteralPath $generatedJson)) {
    throw "Missing generated request JSON: $generatedJson"
}

& $mvn -q '-Dtest=com.lz.paperword.tools.ReferenceRoundTripDocxTest' test
if ($LASTEXITCODE -ne 0) {
    throw "Reference Word regeneration test failed with exit code $LASTEXITCODE"
}
if (-not (Test-Path -LiteralPath $generatedDocx)) {
    throw "Missing regenerated Word document: $generatedDocx"
}

python rebuild\extract_formula_boxes.py $referencePath --out $referenceFormulaBoxes
if ($LASTEXITCODE -ne 0) {
    throw "Reference formula box extraction failed with exit code $LASTEXITCODE"
}
python rebuild\extract_formula_boxes.py $generatedDocx --out $uncalibratedFormulaBoxes
if ($LASTEXITCODE -ne 0) {
    throw "Uncalibrated formula box extraction failed with exit code $LASTEXITCODE"
}

if (-not $SkipFormulaDataset) {
    python rebuild\build_formula_box_dataset.py --source-root $FormulaDatasetRoot --out-json $formulaDatasetJson --out-text $formulaDatasetText --max-docs $FormulaDatasetMaxDocs --min-formulas 10
    if ($LASTEXITCODE -ne 0) {
        throw "Formula box dataset build failed with exit code $LASTEXITCODE"
    }
}

python rebuild\fit_formula_box_model.py --reference $referencePath --generated $generatedDocx --reference-boxes $referenceFormulaBoxes --generated-boxes $uncalibratedFormulaBoxes --request-json $generatedJson --dataset-json $formulaDatasetJson --out-json $formulaFitJson --out-text $formulaFitText
if ($LASTEXITCODE -ne 0) {
    throw "Formula box fitting failed with exit code $LASTEXITCODE"
}

python rebuild\calibrate_formula_boxes.py --reference $referencePath --input $generatedDocx --output $generatedDocx
if ($LASTEXITCODE -ne 0) {
    throw "Formula box calibration failed with exit code $LASTEXITCODE"
}

python rebuild\sync_formula_vector_previews.py --reference $referencePath --input $generatedDocx --output $generatedDocx --report $formulaVectorSyncJson
if ($LASTEXITCODE -ne 0) {
    throw "Formula vector preview synchronization failed with exit code $LASTEXITCODE"
}

python rebuild\extract_formula_boxes.py $generatedDocx --out $calibratedFormulaBoxes
if ($LASTEXITCODE -ne 0) {
    throw "Calibrated formula box extraction failed with exit code $LASTEXITCODE"
}

.\scripts\verify-mathtype-word.ps1 -DocxPath $generatedDocx -MinimumOleCount $MinimumOleCount
$docx2texOutDir = Join-Path $OutDir "regenerated-docx2tex"
.\scripts\verify-docx2tex-roundtrip.ps1 `
    -DocxPath $generatedDocx `
    -Docx2TexRoot $Docx2TexRoot `
    -OutDir $docx2texOutDir `
    -RequestJsonPath $generatedJson `
    -MinimumTimesCount 1000 `
    -MinimumCdotsCount 150 `
    -MinimumFractionCount 1500 `
    -ExpectedFragments @(
        "\frac{99}{1\times 2\times 3\times \cdots \times 100}",
        "\frac{1}{1\times 2}-\frac{1}{1\times 2\times 3\times \cdots \times 100}",
        "\frac{2}{1\times (1+2)}+\frac{3}{(1+2)\times (1+2+3)}"
    )
$docx2texTex = Join-Path $docx2texOutDir ([IO.Path]::GetFileNameWithoutExtension($generatedDocx) + ".tex")
python rebuild\verify_docx2tex_formula_fragments.py `
    --request-json $generatedJson `
    --tex $docx2texTex `
    --out-json (Join-Path $OutDir "docx2tex-formula-fragment-coverage.json") `
    --out-text (Join-Path $OutDir "docx2tex-formula-fragment-coverage.txt") `
    --require-risk-match `
    --min-risk-coverage 0.98 `
    --require-occurrence-match `
    --min-occurrence-coverage 0.98
if ($LASTEXITCODE -ne 0) {
    throw "docx2tex formula fragment coverage failed with exit code $LASTEXITCODE"
}

if ($SemanticOnly) {
    Write-Output "reference semantic roundtrip verification ok"
    return
}

if ($UseReferenceLayoutShell) {
    python rebuild\sync_reference_layout_shell.py `
        --reference $referencePath `
        --generated $generatedDocx `
        --output $shellGeneratedDocx `
        --report $referenceLayoutShellJson
    if ($LASTEXITCODE -ne 0) {
        throw "Reference layout shell synchronization failed with exit code $LASTEXITCODE"
    }
    $generatedDocx = $shellGeneratedDocx
    .\scripts\verify-mathtype-word.ps1 -DocxPath $generatedDocx -MinimumOleCount $MinimumOleCount
    $shellDocx2texOutDir = Join-Path $OutDir "shell-docx2tex"
    .\scripts\verify-docx2tex-roundtrip.ps1 `
        -DocxPath $generatedDocx `
        -Docx2TexRoot $Docx2TexRoot `
        -OutDir $shellDocx2texOutDir `
        -RequestJsonPath $generatedJson `
        -MinimumTimesCount 1000 `
        -MinimumCdotsCount 150 `
        -MinimumFractionCount 1500 `
        -ExpectedFragments @(
            "\frac{99}{1\times 2\times 3\times \cdots \times 100}",
            "\frac{1}{1\times 2}-\frac{1}{1\times 2\times 3\times \cdots \times 100}",
            "\frac{2}{1\times (1+2)}+\frac{3}{(1+2)\times (1+2+3)}"
        )
    $shellDocx2texTex = Join-Path $shellDocx2texOutDir ([IO.Path]::GetFileNameWithoutExtension($generatedDocx) + ".tex")
    python rebuild\verify_docx2tex_formula_fragments.py `
        --request-json $generatedJson `
        --tex $shellDocx2texTex `
        --out-json (Join-Path $OutDir "shell-docx2tex-formula-fragment-coverage.json") `
        --out-text (Join-Path $OutDir "shell-docx2tex-formula-fragment-coverage.txt") `
        --require-risk-match `
        --min-risk-coverage 0.98 `
        --require-occurrence-match `
        --min-occurrence-coverage 0.98
    if ($LASTEXITCODE -ne 0) {
        throw "reference layout shell docx2tex formula fragment coverage failed with exit code $LASTEXITCODE"
    }
}

$metrics = @(
    Get-DocxMetrics $referencePath
    Get-DocxMetrics $generatedDocx
)
$referenceMetrics = $metrics[0]
$generatedMetrics = $metrics[1]
if ($generatedMetrics.ShapeHeightAvgPt -lt ($referenceMetrics.ShapeHeightAvgPt * $MinimumFormulaHeightRatio)) {
    throw "Generated formula average height is too small: generated=$($generatedMetrics.ShapeHeightAvgPt)pt reference=$($referenceMetrics.ShapeHeightAvgPt)pt ratioMin=$MinimumFormulaHeightRatio"
}
if ($generatedMetrics.ShapeWidthMaxPt -gt ($referenceMetrics.ShapeWidthMaxPt + $MaximumFormulaWidthOverflowPt)) {
    throw "Generated formula max width is too large: generated=$($generatedMetrics.ShapeWidthMaxPt)pt reference=$($referenceMetrics.ShapeWidthMaxPt)pt overflowMax=$MaximumFormulaWidthOverflowPt"
}
if (-not $generatedMetrics.MediaExt.Contains("wmf=")) {
    throw "Generated formula previews are not vector WMF previews: media=$($generatedMetrics.MediaExt)"
}
$metrics | ConvertTo-Json -Depth 4

python rebuild\compare_reference_format.py --reference $referencePath --generated $generatedDocx --request-json $generatedJson --out-json (Join-Path $OutDir "format-comparison.json") --out-text (Join-Path $OutDir "format-comparison.txt")
if ($LASTEXITCODE -ne 0) {
    throw "Reference format comparison failed with exit code $LASTEXITCODE"
}

.\rebuild\compare_word_formula_layout.ps1 -ReferenceDocx $referencePath -GeneratedDocx $generatedDocx -OutJson $wordFormulaLayoutJson -OutText $wordFormulaLayoutText
if ($LASTEXITCODE -ne 0) {
    throw "Word formula layout comparison failed with exit code $LASTEXITCODE"
}

if (-not $SkipVisual) {
    Export-DocxPdf $referencePath $referencePdf
    Export-DocxPdf $generatedDocx $generatedPdf
    Export-PdfPages $referencePdf $referencePages
    Export-PdfPages $generatedPdf $generatedPages
    $visualMetrics = Compare-VisualPages $referencePages $generatedPages
    $referencePageCount = @(Get-ChildItem -LiteralPath $referencePages -Filter "*.png").Count
    $generatedPageCount = @(Get-ChildItem -LiteralPath $generatedPages -Filter "*.png").Count
    if ($referencePageCount -ne $generatedPageCount) {
        throw "visual page count mismatch: reference=$referencePageCount generated=$generatedPageCount"
    }
    if ($CheckRenderedFormulaInk) {
        python rebuild\compare_rendered_formula_ink.py `
            --layout-json $wordFormulaLayoutJson `
            --reference-pages $referencePages `
            --generated-pages $generatedPages `
            --out-json $renderedFormulaInkJson `
            --out-text $renderedFormulaInkText `
            --fail
        if ($LASTEXITCODE -ne 0) {
            throw "Rendered formula ink comparison failed with exit code $LASTEXITCODE"
        }
    }
    $visualMetrics | ConvertTo-Json -Depth 4
}

Write-Output "reference roundtrip verification ok"
