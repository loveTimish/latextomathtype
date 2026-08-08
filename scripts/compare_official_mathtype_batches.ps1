param(
    [string]$Manifest = "target\mathtype-standard\manifest.json",
    [string]$Python = "python",
    [string]$Magick = "magick"
)

$ErrorActionPreference = "Stop"
$repoRoot = Resolve-Path (Join-Path $PSScriptRoot "..")
$manifestPath = [IO.Path]::GetFullPath((Join-Path $repoRoot $Manifest))
$data = Get-Content -LiteralPath $manifestPath -Raw -Encoding UTF8 | ConvertFrom-Json
$metricScript = Join-Path $PSScriptRoot "compare_docx_pair_metrics.py"
$visualScript = Join-Path $PSScriptRoot "compare_mathtype_standard_visuals.py"
$batches = @()
$total = 0
$sizeExact = 0
$baselineExact = 0
$visualPassed = 0
$failures = @()

foreach ($batch in $data.batches) {
    $standard = [string]$batch.standardDocx
    $generated = Join-Path (Split-Path -Parent $standard) "generated-formatted.docx"
    if (-not (Test-Path -LiteralPath $generated)) {
        throw "missing generated batch DOCX: $generated"
    }
    $batchDir = Split-Path -Parent $standard
    $formulaItems = @(Get-Content -LiteralPath ([string]$batch.formulas) -Raw -Encoding UTF8 | ConvertFrom-Json)
    $metricsDir = Join-Path $batchDir "metrics"
    $visualReport = Join-Path $batchDir "visual-comparison.json"
    & $Python $metricScript $standard $generated $metricsDir | Out-Null
    if ($LASTEXITCODE -ne 0) { throw "metric comparison failed for $($batch.name)" }
    & $Python $visualScript $standard $generated $visualReport --magick $Magick --no-fail | Out-Null
    if ($LASTEXITCODE -ne 0) { throw "visual comparison failed for $($batch.name)" }

    $metricSummaryPath = Get-ChildItem -LiteralPath $metricsDir -Filter "*_summary.json" | Select-Object -First 1
    $pairCsvPath = Get-ChildItem -LiteralPath $metricsDir -Filter "*.csv" | Select-Object -First 1
    $metricSummary = Get-Content -LiteralPath $metricSummaryPath.FullName -Raw -Encoding UTF8 | ConvertFrom-Json
    $visual = Get-Content -LiteralPath $visualReport -Raw -Encoding UTF8 | ConvertFrom-Json
    $pairs = @(Import-Csv -LiteralPath $pairCsvPath.FullName)
    if ($pairs.Count -ne $formulaItems.Count) {
        $failures += [pscustomobject]@{
            batch = [string]$batch.name
            index = -1
            globalFormulaIndex = -1
            formula = ""
            code = "OBJECT_PAIR_COUNT_MISMATCH"
            expected = $formulaItems.Count
            actual = $pairs.Count
        }
    }
    foreach ($pair in $pairs) {
        $total++
        $formulaIndex = [int]$pair.index - 1
        $formulaItem = if ($formulaIndex -ge 0 -and $formulaIndex -lt $formulaItems.Count) {
            $formulaItems[$formulaIndex]
        } else { $null }
        $sizeFieldPairs = @(
            @($pair.source_shape_w_pt, $pair.generated_shape_w_pt, $pair.shape_w_pt_error_pct),
            @($pair.source_shape_h_pt, $pair.generated_shape_h_pt, $pair.shape_h_pt_error_pct),
            @($pair.source_wmf_w_pt, $pair.generated_wmf_w_pt, $pair.wmf_w_pt_error_pct),
            @($pair.source_wmf_h_pt, $pair.generated_wmf_h_pt, $pair.wmf_h_pt_error_pct),
            @($pair.source_dxa_pt, $pair.generated_dxa_pt, $pair.dxa_pt_error_pct),
            @($pair.source_dya_pt, $pair.generated_dya_pt, $pair.dya_pt_error_pct)
        )
        $exactSize = @($sizeFieldPairs | Where-Object {
            $sourceMissing = [string]::IsNullOrWhiteSpace([string]$_[0])
            $generatedMissing = [string]::IsNullOrWhiteSpace([string]$_[1])
            ($sourceMissing -xor $generatedMissing) `
                -or (-not $sourceMissing -and [double]$_[2] -ne 0.0)
        }).Count -eq 0
        $exactBaseline = [string]$pair.source_position_halfpt -eq [string]$pair.generated_position_halfpt
        if ($exactSize) { $sizeExact++ }
        if ($exactBaseline) { $baselineExact++ }
        if (-not $exactSize -or -not $exactBaseline) {
            $failures += [pscustomobject]@{
                batch = [string]$batch.name
                index = [int]$pair.index
                globalFormulaIndex = if ($formulaItem) { [int]$formulaItem.globalIndex } else { -1 }
                formula = if ($formulaItem) { [string]$formulaItem.formula } else { "" }
                code = if (-not $exactSize) { "SIZE_MISMATCH" } else { "BASELINE_MISMATCH" }
                shapeWidthErrorPct = [double]$pair.shape_w_pt_error_pct
                shapeHeightErrorPct = [double]$pair.shape_h_pt_error_pct
                wmfWidthErrorPct = [double]$pair.wmf_w_pt_error_pct
                wmfHeightErrorPct = [double]$pair.wmf_h_pt_error_pct
                baselineDiffHalfPt = [int]$pair.position_diff_halfpt
            }
        }
    }
    $visualPassed += [int]$visual.passedCount
    foreach ($formula in $visual.formulas) {
        if (-not $formula.passed) {
            $formulaItem = if ([int]$formula.index -ge 0 -and [int]$formula.index -lt $formulaItems.Count) {
                $formulaItems[[int]$formula.index]
            } else { $null }
            $failures += [pscustomobject]@{
                batch = [string]$batch.name
                index = [int]$formula.index + 1
                globalFormulaIndex = if ($formulaItem) { [int]$formulaItem.globalIndex } else { -1 }
                formula = if ($formulaItem) { [string]$formulaItem.formula } else { "" }
                code = "VISUAL_MISMATCH"
                ssim = $formula.ssim
                foregroundIou = $formula.foregroundIou
                missingMajorComponents = $formula.missingMajorComponents
                error = $formula.error
            }
        }
    }
    $batches += [pscustomobject]@{
        name = [string]$batch.name
        formulaCount = [int]$batch.formulaCount
        metrics = $metricSummaryPath.FullName
        visuals = $visualReport
    }
}

$report = [ordered]@{
    schemaVersion = 1
    officialEntryCount = [int]$data.officialEntryCount
    selectedFormulaCount = [int]$data.selectedFormulaCount
    pairedObjectCount = $total
    exactSizeCount = $sizeExact
    exactBaselineCount = $baselineExact
    visualPassedCount = $visualPassed
    thresholds = [ordered]@{ ssim = 0.90; foregroundIou = 0.85; missingMajorComponents = 0 }
    batches = $batches
    failureCount = $failures.Count
    failures = $failures
}
$reportPath = Join-Path (Split-Path -Parent $manifestPath) "final-layout-visual-report.json"
$report | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $reportPath -Encoding UTF8
Write-Host "report=$reportPath"
Write-Host "paired=$total exactSize=$sizeExact exactBaseline=$baselineExact visualPassed=$visualPassed"
if ($total -ne [int]$data.selectedFormulaCount -or $sizeExact -ne $total `
        -or $baselineExact -ne $total -or $visualPassed -ne $total -or $failures.Count -ne 0) {
    throw "MathType layout/visual gate failed; inspect $reportPath"
}
