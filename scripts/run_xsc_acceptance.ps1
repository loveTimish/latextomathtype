param(
    [int]$Start = 1,
    [int]$End = 10,
    [string]$Stamp = "",
    [string]$DatasetDir = "F:\资料\xsc资料\word_files",
    [string]$AnalysisDir = "D:\latextomathtype\analysis",
    [string]$LatexRoot = "D:\latextomathtype\analysis\xsc-latex",
    [string]$DocxToLatexDir = "D:\docxtolatex\docxtolatex",
    [switch]$SkipGenerate
)

$ErrorActionPreference = "Stop"

if (-not $Stamp) {
    $Stamp = Get-Date -Format "yyyyMMdd-HHmmss"
}

$repoRoot = Resolve-Path (Join-Path $PSScriptRoot "..")
$docxOut = Join-Path $AnalysisDir "batch10-full-docx\$Stamp"
$sizeOut = Join-Path $AnalysisDir "pair-metrics-latex\$Stamp-target-direct"
$mtefOut = Join-Path $AnalysisDir "mtef-report\$Stamp-keyed-object"

Write-Host "stamp=$Stamp"
Write-Host "range=$Start..$End"

function Invoke-Checked {
    param(
        [string]$Tool,
        [string[]]$Arguments,
        [switch]$Quiet
    )
    if ($Quiet) {
        & $Tool @Arguments | Out-Null
    } else {
        & $Tool @Arguments
    }
    if ($LASTEXITCODE -ne 0) {
        throw "$Tool failed with exit code $LASTEXITCODE"
    }
}

if (-not $SkipGenerate) {
    Push-Location $repoRoot
    try {
        mvn -q -Dtest=XscFullBatch10DocxTest "-Dxsc.full.start=$Start" "-Dxsc.full.end=$End" test
    } finally {
        Pop-Location
    }
    $latest = Get-ChildItem (Join-Path $AnalysisDir "batch10-full-docx") -Directory |
        Sort-Object Name -Descending |
        Select-Object -First 1
    if ($latest -and $latest.Name -ne $Stamp) {
        $docxOut = $latest.FullName
        $Stamp = $latest.Name
        $sizeOut = Join-Path $AnalysisDir "pair-metrics-latex\$Stamp-target-direct"
        $mtefOut = Join-Path $AnalysisDir "mtef-report\$Stamp-keyed-object"
        Write-Host "generatedStamp=$Stamp"
    }
}

New-Item -ItemType Directory -Force -Path $sizeOut | Out-Null
for ($i = $Start; $i -le $End; $i++) {
    $pad = "{0:D2}" -f $i
    $sourceDocx = Join-Path $DatasetDir "$i.docx"
    $generatedDocx = Join-Path $docxOut "xsc测试集完整重建_$pad.docx"
    $reportPath = Join-Path $LatexRoot "$i\$i.report.json"
    if (-not (Test-Path -LiteralPath $reportPath)) {
        $fallbackReport = Join-Path $AnalysisDir "batch10-latex\$i\$i.report.json"
        if (Test-Path -LiteralPath $fallbackReport) {
            $reportPath = $fallbackReport
        }
    }
    if (-not (Test-Path -LiteralPath $sourceDocx)) { throw "missing source docx: $sourceDocx" }
    if (-not (Test-Path -LiteralPath $generatedDocx)) { throw "missing generated docx: $generatedDocx" }
    if (-not (Test-Path -LiteralPath $reportPath)) { throw "missing report json: $reportPath" }
    Invoke-Checked "python" @(
        (Join-Path $AnalysisDir "compare_docx_pair_metrics.py"),
        $sourceDocx,
        $generatedDocx,
        $sizeOut,
        $reportPath,
        (Join-Path $AnalysisDir "batch10-full-requests\full-$pad.request.json")
    ) -Quiet
}

New-Item -ItemType Directory -Force -Path $mtefOut | Out-Null
Push-Location $DocxToLatexDir
try {
    for ($i = $Start; $i -le $End; $i++) {
        $pad = "{0:D2}" -f $i
        $reportPath = Join-Path $LatexRoot "$i\$i.report.json"
        if (-not (Test-Path -LiteralPath $reportPath)) {
            $fallbackReport = Join-Path $AnalysisDir "batch10-latex\$i\$i.report.json"
            if (Test-Path -LiteralPath $fallbackReport) {
                $reportPath = $fallbackReport
            }
        }
        if (-not (Test-Path -LiteralPath $reportPath)) { throw "missing report json: $reportPath" }
        Invoke-Checked "go" @(
            "run",
            (Join-Path $AnalysisDir "compare_mtef_native.go"),
            (Join-Path $DatasetDir "$i.docx"),
            (Join-Path $docxOut "xsc测试集完整重建_$pad.docx"),
            (Join-Path $mtefOut "$i"),
            $reportPath,
            (Join-Path $AnalysisDir "batch10-full-requests\full-$pad.request.json")
        ) -Quiet
    }
} finally {
    Pop-Location
}

Invoke-Checked "python" @(
    (Join-Path $AnalysisDir "summarize_xsc_acceptance.py"),
    $Stamp,
    $sizeOut,
    $mtefOut,
    "--start",
    "$Start",
    "--end",
    "$End"
)
