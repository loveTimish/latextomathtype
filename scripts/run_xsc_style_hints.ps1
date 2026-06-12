param(
    [int]$Start = 1,
    [int]$End = 155,
    [string]$DatasetDir = "F:\资料\xsc资料\word_files",
    [string]$AnalysisDir = "D:\latextomathtype\analysis",
    [string]$DocxToLatexDir = "D:\docxtolatex\docxtolatex"
)

$ErrorActionPreference = "Stop"

$repoRoot = Resolve-Path (Join-Path $PSScriptRoot "..")
$outDir = Join-Path $AnalysisDir "mtef-style-hints"
New-Item -ItemType Directory -Force -Path $outDir | Out-Null

Push-Location $DocxToLatexDir
try {
    for ($i = $Start; $i -le $End; $i++) {
        $sourceDocx = Join-Path $DatasetDir "$i.docx"
        $outPath = Join-Path $outDir "$i.style-hints.json"
        if (-not (Test-Path -LiteralPath $sourceDocx)) {
            throw "missing source docx: $sourceDocx"
        }
        go run (Join-Path $repoRoot "scripts\extract_mtef_style_hints.go") $sourceDocx $outPath
        if ($LASTEXITCODE -ne 0) {
            throw "extract_mtef_style_hints failed for $sourceDocx"
        }
    }
} finally {
    Pop-Location
}

Write-Host "styleHints=$outDir"
