param(
    [int]$Start = 1,
    [int]$End = 155,
    [string]$DatasetDir = "F:\资料\xsc资料\word_files",
    [string]$OutRoot = "J:\latextomathtype\analysis\xsc-latex-fixed",
    [string]$DocxToLatexDir = "J:\docxtolatex\docxtolatex-main"
)

$ErrorActionPreference = "Stop"

function Invoke-Checked {
    param(
        [string]$Tool,
        [string[]]$Arguments
    )
    & $Tool @Arguments
    if ($LASTEXITCODE -ne 0) {
        throw "$Tool failed with exit code $LASTEXITCODE"
    }
}

New-Item -ItemType Directory -Force -Path $OutRoot | Out-Null
if (-not (Test-Path -LiteralPath $DocxToLatexDir)) {
    throw "missing docxtolatex dir: $DocxToLatexDir"
}

Push-Location $DocxToLatexDir
try {
    for ($i = $Start; $i -le $End; $i++) {
        $sourceDocx = Join-Path $DatasetDir "$i.docx"
        if (-not (Test-Path -LiteralPath $sourceDocx)) {
            throw "missing source docx: $sourceDocx"
        }
        $outDir = Join-Path $OutRoot "$i"
        New-Item -ItemType Directory -Force -Path $outDir | Out-Null
        Write-Host "docxtolatex $i -> $outDir"
        Invoke-Checked "go" @(
            "run",
            "main.go",
            "--wordDocx",
            $sourceDocx,
            "--output",
            $outDir,
            "--report"
        )
    }
} finally {
    Pop-Location
}
