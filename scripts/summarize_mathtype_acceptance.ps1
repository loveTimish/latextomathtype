param(
    [string]$Out = "target\mathtype-acceptance\current-summary.json",
    [string]$OfficialReport = "target\official-latex-coverage\current-report.json",
    [string]$CombinationReport = "target\official-latex-coverage\three-level-combinations.json",
    [string]$SymbolReport = "target\official-latex-coverage\symbol-encoding-report.json",
    [string]$StandardStructureReport = "target\mathtype-standard\final-formatted-structure-report.json",
    [string]$StandardLayoutVisualReport = "target\mathtype-standard\final-layout-visual-report.json",
    [string]$XscReport = "analysis\acceptance-summary\xsc-full-acceptance.json",
    [string]$TenthFormulaReport = "target\mathtype-tex-variants\arrow-final-raw-structure.json"
)

$ErrorActionPreference = "Stop"
$repoRoot = Resolve-Path (Join-Path $PSScriptRoot "..")
function Resolve-RepoPath([string]$Path) {
    if ([IO.Path]::IsPathRooted($Path)) { return [IO.Path]::GetFullPath($Path) }
    return [IO.Path]::GetFullPath((Join-Path $repoRoot $Path))
}
function Read-JsonIfPresent([string]$Path) {
    $resolved = Resolve-RepoPath $Path
    if (-not (Test-Path -LiteralPath $resolved)) { return $null }
    return Get-Content -LiteralPath $resolved -Raw -Encoding UTF8 | ConvertFrom-Json
}

$official = Read-JsonIfPresent $OfficialReport
$combinations = Read-JsonIfPresent $CombinationReport
$symbols = Read-JsonIfPresent $SymbolReport
$standardStructure = Read-JsonIfPresent $StandardStructureReport
$standardLayout = Read-JsonIfPresent $StandardLayoutVisualReport
$xsc = Read-JsonIfPresent $XscReport
$tenth = Read-JsonIfPresent $TenthFormulaReport

$testCount = $failureCount = $errorCount = $skippedCount = 0
$surefireDir = Resolve-RepoPath "target\surefire-reports"
if (Test-Path -LiteralPath $surefireDir) {
    foreach ($file in Get-ChildItem -LiteralPath $surefireDir -Filter "TEST-*.xml") {
        [xml]$xml = Get-Content -LiteralPath $file.FullName -Raw
        $testCount += [int]$xml.testsuite.tests
        $failureCount += [int]$xml.testsuite.failures
        $errorCount += [int]$xml.testsuite.errors
        $skippedCount += [int]$xml.testsuite.skipped
    }
}

$verifiedSymbols = if ($symbols) { @($symbols.profiles | Where-Object mathTypeVerified).Count } else { 0 }
$mappedSymbols = if ($symbols) { [int]$symbols.mappedOfficialSymbols } else { 0 }
$gates = [ordered]@{
    mavenTests = $testCount -gt 0 -and $failureCount -eq 0 -and $errorCount -eq 0
    officialCoverage = $official -and [int]$official.totalEntries -eq 551 `
        -and [int]$official.successfulEntries -eq 551 -and [int]$official.failedEntries -eq 0
    threeLevelCombinations = $combinations -and [int]$combinations.generatedCombinationCount -gt 0 `
        -and [int]$combinations.generatedCombinationCount -eq [int]$combinations.normalizedStructureCount `
        -and [int]$combinations.failureCount -eq 0
    allMappedSymbolsMathTypeVerified = $mappedSymbols -gt 0 -and $verifiedSymbols -eq $mappedSymbols
    tenthFormulaExactStructure = $tenth -and [int]$tenth.standardObjectCount -eq 1 `
        -and [int]$tenth.validGeneratedOleCount -eq 1 `
        -and [int]$tenth.normalizedStructureEqualCount -eq 1
    officialStandardStructure = $standardStructure -and [int]$standardStructure.expectedObjectCount -gt 0 `
        -and [int]$standardStructure.expectedObjectCount -eq [int]$standardStructure.normalizedStructureEqualCount
    officialExactLayoutAndVisual = $standardLayout -and [int]$standardLayout.pairedObjectCount -gt 0 `
        -and [int]$standardLayout.pairedObjectCount -eq [int]$standardLayout.exactSizeCount `
        -and [int]$standardLayout.pairedObjectCount -eq [int]$standardLayout.exactBaselineCount `
        -and [int]$standardLayout.pairedObjectCount -eq [int]$standardLayout.visualPassedCount
    xsc39551 = $xsc -and [bool]$xsc.passed `
        -and [int]$xsc.exactMtef.expectedObjects -eq 39551 `
        -and [int]$xsc.exactMtef.normalizedStructureEqual -eq 39551
}
$passed = @($gates.Values | Where-Object { -not $_ }).Count -eq 0

$report = [ordered]@{
    schemaVersion = 1
    generatedAt = (Get-Date).ToString("o")
    passed = $passed
    gates = $gates
    maven = [ordered]@{ tests = $testCount; failures = $failureCount; errors = $errorCount; skipped = $skippedCount }
    official = if ($official) { [ordered]@{
        entries = [int]$official.totalEntries
        successful = [int]$official.successfulEntries
        failed = [int]$official.failedEntries
    } } else { $null }
    combinations = if ($combinations) { [ordered]@{
        generated = [int]$combinations.generatedCombinationCount
        normalized = [int]$combinations.normalizedStructureCount
        excluded = [int]$combinations.exclusionCount
        failed = [int]$combinations.failureCount
    } } else { $null }
    symbols = [ordered]@{ mapped = $mappedSymbols; mathTypeVerified = $verifiedSymbols; unverified = $mappedSymbols - $verifiedSymbols }
    pending = @($gates.Keys | Where-Object { -not $gates[$_] })
}
$outPath = Resolve-RepoPath $Out
New-Item -ItemType Directory -Force -Path (Split-Path -Parent $outPath) | Out-Null
$report | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $outPath -Encoding UTF8
Write-Host "report=$outPath"
Write-Host "passed=$passed"
if (-not $passed) { exit 1 }
