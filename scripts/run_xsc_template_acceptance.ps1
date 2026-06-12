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
$generatedDir = Join-Path $AnalysisDir "batch10-full-docx\$Stamp"
$templateDir = Join-Path $AnalysisDir "template-rebuild\$Stamp"
$pairDir = Join-Path $templateDir "mtef-pairs"
$recordDir = Join-Path $templateDir "wmf-records"
$metricDir = Join-Path $templateDir "metrics"

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
        Invoke-Checked "mvn" @(
            "-q",
            "-Dtest=XscFullBatch10DocxTest#generateTenFullXscDocxFiles",
            "-Dxsc.full.start=$Start",
            "-Dxsc.full.end=$End",
            "test"
        )
    } finally {
        Pop-Location
    }
    $latest = Get-ChildItem (Join-Path $AnalysisDir "batch10-full-docx") -Directory |
        Sort-Object Name -Descending |
        Select-Object -First 1
    if ($latest) {
        $generatedDir = $latest.FullName
        $Stamp = $latest.Name
        $templateDir = Join-Path $AnalysisDir "template-rebuild\$Stamp"
        $pairDir = Join-Path $templateDir "mtef-pairs"
        $recordDir = Join-Path $templateDir "wmf-records"
        $metricDir = Join-Path $templateDir "metrics"
    }
}

New-Item -ItemType Directory -Force -Path $templateDir, $pairDir, $recordDir, $metricDir | Out-Null

$items = @()
Push-Location $DocxToLatexDir
try {
    for ($i = $Start; $i -le $End; $i++) {
        $pad = "{0:D2}" -f $i
        $sourceDocx = Join-Path $DatasetDir "$i.docx"
        $generatedDocx = Join-Path $generatedDir "xsc测试集完整重建_$pad.docx"
        $reportPath = Join-Path $LatexRoot "$i\$i.report.json"
        if (-not (Test-Path -LiteralPath $reportPath)) {
            $fallbackReport = Join-Path $AnalysisDir "batch10-latex\$i\$i.report.json"
            if (Test-Path -LiteralPath $fallbackReport) {
                $reportPath = $fallbackReport
            }
        }
        $requestPath = Join-Path $AnalysisDir "batch10-full-requests\full-$pad.request.json"
        $pairOut = Join-Path $pairDir "$i"
        $templateDocx = Join-Path $templateDir "xsc模板保真重建_$pad.docx"
        $recordJson = Join-Path $recordDir "$i-wmf-records.json"
        $metricOut = Join-Path $metricDir "$i"
        New-Item -ItemType Directory -Force -Path $pairOut, $metricOut | Out-Null

        if (-not (Test-Path -LiteralPath $sourceDocx)) { throw "missing source docx: $sourceDocx" }
        if (-not (Test-Path -LiteralPath $generatedDocx)) { throw "missing generated docx: $generatedDocx" }
        if (-not (Test-Path -LiteralPath $reportPath)) { throw "missing report json: $reportPath" }
        if (-not (Test-Path -LiteralPath $requestPath)) { throw "missing request json: $requestPath" }

        Invoke-Checked "go" @(
            "run",
            (Join-Path $AnalysisDir "compare_mtef_native.go"),
            $sourceDocx,
            $generatedDocx,
            $pairOut,
            $reportPath,
            $requestPath
        ) -Quiet
        $pairCsv = Join-Path $pairOut "$i`_mtef_pairs.csv"
        if (-not (Test-Path -LiteralPath $pairCsv)) { throw "missing pair csv: $pairCsv" }

        Invoke-Checked "python" @(
            (Join-Path $repoRoot "scripts\rebuild_docx_from_source_template.py"),
            $sourceDocx,
            $generatedDocx,
            $pairCsv,
            $templateDocx
        ) -Quiet

        Invoke-Checked "python" @(
            (Join-Path $repoRoot "scripts\wmf_record_report.py"),
            $templateDocx,
            "--out",
            $recordJson
        ) -Quiet

        Invoke-Checked "python" @(
            (Join-Path $AnalysisDir "compare_docx_pair_metrics.py"),
            $sourceDocx,
            $templateDocx,
            $metricOut
        ) -Quiet

        $record = Get-Content -LiteralPath $recordJson -Raw | ConvertFrom-Json
        $metricSummary = Get-ChildItem -LiteralPath $metricOut -Filter "*_summary.json" | Select-Object -First 1
        $metric = Get-Content -LiteralPath $metricSummary.FullName -Raw | ConvertFrom-Json
        $items += [pscustomobject]@{
            doc = $i
            source = $sourceDocx
            generated = $generatedDocx
            template = $templateDocx
            pairCsv = $pairCsv
            wmf = [int]$record.wmf
            vectorText = [int]$record.vectorText
            stretchDib = [int]$record.stretchDib
            pairedObjects = [int]$metric.paired_objects
            nonWmfGenerated = [int]$metric.non_wmf_generated
            wmfWidthWithin1pct = [int]$metric.wmf_width_ratio.within_1pct
            wmfHeightWithin1pct = [int]$metric.wmf_height_ratio.within_1pct
            wmfWidthMaxErrorPct = [double]$metric.wmf_width_ratio.max_abs_error_pct
            wmfHeightMaxErrorPct = [double]$metric.wmf_height_ratio.max_abs_error_pct
            shapeWidthMaxErrorPct = [double]$metric.shape_width_ratio.max_abs_error_pct
            shapeHeightMaxErrorPct = [double]$metric.shape_height_ratio.max_abs_error_pct
        }
    }
} finally {
    Pop-Location
}

$summary = [pscustomobject]@{
    stamp = $Stamp
    range = @{ start = $Start; end = $End }
    templateDir = $templateDir
    docs = $items.Count
    totals = @{
        wmf = ($items | Measure-Object -Property wmf -Sum).Sum
        vectorText = ($items | Measure-Object -Property vectorText -Sum).Sum
        stretchDib = ($items | Measure-Object -Property stretchDib -Sum).Sum
        pairedObjects = ($items | Measure-Object -Property pairedObjects -Sum).Sum
        nonWmfGenerated = ($items | Measure-Object -Property nonWmfGenerated -Sum).Sum
    }
    worst = @{
        wmfWidthMaxErrorPct = ($items | Measure-Object -Property wmfWidthMaxErrorPct -Maximum).Maximum
        wmfHeightMaxErrorPct = ($items | Measure-Object -Property wmfHeightMaxErrorPct -Maximum).Maximum
        shapeWidthMaxErrorPct = ($items | Measure-Object -Property shapeWidthMaxErrorPct -Maximum).Maximum
        shapeHeightMaxErrorPct = ($items | Measure-Object -Property shapeHeightMaxErrorPct -Maximum).Maximum
    }
    items = $items
}

$summaryPath = Join-Path $templateDir "template-acceptance-summary.json"
$summary | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $summaryPath -Encoding UTF8
Write-Host $summaryPath
