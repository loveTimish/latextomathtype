param(
    [string]$Manifest = "target\mathtype-standard\manifest.json"
)

$ErrorActionPreference = "Stop"
$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$manifestPath = if ([IO.Path]::IsPathRooted($Manifest)) {
    [IO.Path]::GetFullPath($Manifest)
} else {
    [IO.Path]::GetFullPath((Join-Path $repoRoot $Manifest))
}
$data = Get-Content -LiteralPath $manifestPath -Raw -Encoding UTF8 | ConvertFrom-Json
$classpathPath = Join-Path $repoRoot "target\classpath.txt"
if (-not (Test-Path -LiteralPath $classpathPath)) {
    throw "missing Maven runtime classpath: $classpathPath"
}
$javaClasspath = (Join-Path $repoRoot "target\classes") + ";" +
    (Get-Content -LiteralPath $classpathPath -Raw).Trim()
$reportRoot = Join-Path (Split-Path -Parent $manifestPath) "structure-comparison"
New-Item -ItemType Directory -Force -Path $reportRoot | Out-Null

$standardObjects = 0
$generatedObjects = 0
$validOle = 0
$structureEqual = 0
$rawEqual = 0
$failures = @()
$batches = @()

foreach ($batch in $data.batches) {
    $standard = [string]$batch.standardDocx
    $generated = Join-Path (Split-Path -Parent $standard) "generated-formatted.docx"
    if (-not (Test-Path -LiteralPath $generated)) {
        throw "missing formatted project DOCX: $generated"
    }
    $batchReport = Join-Path $reportRoot ("{0}.json" -f [string]$batch.name)
    & java -cp $javaClasspath com.lz.paperword.core.mtef.MathTypeDocxComparatorCli `
        $standard $generated $batchReport | Out-Null
    if ($LASTEXITCODE -gt 1 -or -not (Test-Path -LiteralPath $batchReport)) {
        throw "MTEF comparison failed for $($batch.name)"
    }
    $comparison = Get-Content -LiteralPath $batchReport -Raw -Encoding UTF8 | ConvertFrom-Json
    $standardObjects += [int]$comparison.standardObjectCount
    $generatedObjects += [int]$comparison.generatedObjectCount
    $validOle += [int]$comparison.validGeneratedOleCount
    $structureEqual += [int]$comparison.normalizedStructureEqualCount
    $rawEqual += [int]$comparison.rawMtefEqualCount
    foreach ($formula in $comparison.formulas) {
        if (-not [bool]$formula.oleValid -or -not [bool]$formula.normalizedStructureEqual) {
            $failures += [pscustomobject]@{
                batch = [string]$batch.name
                index = [int]$formula.index
                oleValid = [bool]$formula.oleValid
                normalizedStructureEqual = [bool]$formula.normalizedStructureEqual
                standardError = [string]$formula.standardError
                generatedError = [string]$formula.generatedError
                firstNormalizedDifferenceIndex = [int]$formula.firstNormalizedDifferenceIndex
            }
        }
    }
    $batches += [pscustomobject]@{
        name = [string]$batch.name
        formulaCount = [int]$batch.formulaCount
        report = $batchReport
    }
}

$summary = [ordered]@{
    schemaVersion = 1
    officialEntryCount = [int]$data.officialEntryCount
    selectedFormulaCount = [int]$data.selectedFormulaCount
    standardObjectCount = $standardObjects
    generatedObjectCount = $generatedObjects
    validGeneratedOleCount = $validOle
    normalizedStructureEqualCount = $structureEqual
    rawMtefEqualCount = $rawEqual
    failureCount = $failures.Count
    failures = $failures
    batches = $batches
}
$summaryPath = Join-Path (Split-Path -Parent $manifestPath) "final-structure-report.json"
$summary | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $summaryPath -Encoding UTF8
Write-Host "report=$summaryPath"
Write-Host "objects=$generatedObjects validOle=$validOle structureEqual=$structureEqual rawEqual=$rawEqual"

$expected = [int]$data.selectedFormulaCount
if ($standardObjects -ne $expected -or $generatedObjects -ne $expected `
        -or $validOle -ne $expected -or $structureEqual -ne $expected `
        -or $failures.Count -ne 0) {
    throw "MathType normalized structure gate failed; inspect $summaryPath"
}
