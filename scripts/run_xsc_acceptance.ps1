param(
    [int]$Start = 1,
    [int]$End = 10,
    [string]$Stamp = "",
    [string]$DatasetDir = "F:\资料\xsc资料\word_files",
    [string]$AnalysisDir = "J:\latextomathtype\analysis",
    [string]$LatexRoot = "J:\latextomathtype\analysis\xsc-latex-fixed",
    [string]$DocxToLatexDir = "J:\docxtolatex\docxtolatex-main",
    [string]$GeneratedDir = "",
    [string]$RequestDir = "",
    [string]$DatasetManifest = "",
    [string]$FormatPreferenceDocx = "",
    [double]$MaxTargetSizeErrorPct = 1.0,
    [int]$ExpectedOleTotal = 39551,
    [switch]$SkipMtef,
    [switch]$SkipGenerate,
    [switch]$GenerateOnly
)

$ErrorActionPreference = "Stop"

if (-not $Stamp) {
    $Stamp = Get-Date -Format "yyyyMMdd-HHmmss"
}

$repoRoot = Resolve-Path (Join-Path $PSScriptRoot "..")
$mvn = Join-Path $repoRoot ".mvn\apache-maven-3.9.12\bin\mvn.cmd"
if (-not (Test-Path -LiteralPath $mvn)) {
    throw "Bundled Maven not found: $mvn"
}
$formatPreferencePath = if ([string]::IsNullOrWhiteSpace($FormatPreferenceDocx)) {
    $null
} elseif ([IO.Path]::IsPathRooted($FormatPreferenceDocx)) {
    [IO.Path]::GetFullPath($FormatPreferenceDocx)
} else {
    [IO.Path]::GetFullPath((Join-Path $repoRoot $FormatPreferenceDocx))
}
if ($formatPreferencePath -and -not (Test-Path -LiteralPath $formatPreferencePath)) {
    throw "missing pinned MathType format preference DOCX: $formatPreferencePath"
}
$docxOut = if ($GeneratedDir) { $GeneratedDir } else { Join-Path $AnalysisDir "batch10-full-docx\$Stamp" }
$requestOut = if ($RequestDir) { $RequestDir } else { Join-Path $AnalysisDir "batch10-full-requests\$Stamp" }
$styleHintsOut = Join-Path $AnalysisDir "mtef-style-hints"
$sizeOut = Join-Path $AnalysisDir "pair-metrics-latex\$Stamp-target-direct"
$mtefOut = Join-Path $AnalysisDir "mtef-report\$Stamp-keyed-object"
$exactMtefOut = Join-Path $AnalysisDir "mtef-exact-report\$Stamp"
$strictOut = Join-Path $AnalysisDir "acceptance-strict\$Stamp"
$runRoot = Join-Path $AnalysisDir "unattended-runs\$Stamp"
$compareMetricsScript = Join-Path $repoRoot "scripts\compare_docx_pair_metrics.py"
if (-not (Test-Path -LiteralPath $compareMetricsScript)) {
    $compareMetricsScript = Join-Path $AnalysisDir "compare_docx_pair_metrics.py"
}
$compareMtefScript = Join-Path $repoRoot "scripts\compare_mtef_native.go"
if (-not (Test-Path -LiteralPath $compareMtefScript)) {
    throw "missing compare mtef script: $compareMtefScript"
}
$restoreGeometryScript = Join-Path $repoRoot "scripts\restore_formatted_docx_geometry.py"
if (-not (Test-Path -LiteralPath $restoreGeometryScript)) {
    throw "missing formatted geometry restore script: $restoreGeometryScript"
}

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
        $output = & $Tool @Arguments
        if ($output) {
            $output | Write-Output
        }
    }
    if ($LASTEXITCODE -ne 0) {
        throw "$Tool failed with exit code $LASTEXITCODE"
    }
    if (-not $Quiet) {
        return $output
    }
}

function Find-GeneratedDocx {
    param(
        [string]$Directory,
        [string]$Pad
    )
    $matches = @(Get-ChildItem -LiteralPath $Directory -Filter "*_$Pad.docx" -ErrorAction SilentlyContinue)
    if ($matches.Count -eq 1) {
        return $matches[0].FullName
    }
    if ($matches.Count -gt 1) {
        throw "ambiguous generated docx for $Pad in $Directory`: $($matches.Name -join ', ')"
    }
    $fallback = Join-Path $Directory "$Pad.docx"
    if (Test-Path -LiteralPath $fallback) {
        return $fallback
    }
    return Join-Path $Directory "*_$Pad.docx"
}

function Resolve-SourceDocx {
    param(
        [int]$Index
    )
    if ($manifestItems.Count -gt 0) {
        $manifestItem = @($manifestItems | Where-Object { [int]$_.index -eq $Index } | Select-Object -First 1)
        if ($manifestItem) {
            return $manifestItem.sourcePath
        }
    }
    return Join-Path $DatasetDir "$Index.docx"
}

function Read-ManifestItems {
    param(
        [string]$Path
    )
    if (-not $Path) {
        return @()
    }
    if (-not (Test-Path -LiteralPath $Path)) {
        throw "missing dataset manifest: $Path"
    }
    $data = Get-Content -LiteralPath $Path -Raw -Encoding UTF8 | ConvertFrom-Json
    if ($data -isnot [array]) {
        $data = @($data)
    }
    return @($data | Where-Object { $null -ne $_.index -and $_.sourcePath } | ForEach-Object {
        [pscustomobject]@{
            index = [int]$_.index
            sourcePath = [string]$_.sourcePath
            sourceName = if ($_.sourceName) { [string]$_.sourceName } else { "" }
        }
    })
}

$manifestItems = @(Read-ManifestItems $DatasetManifest)

function Read-RequestBuildItem {
    param(
        [string]$RequestDir,
        [int]$Index
    )
    $summaryPath = Join-Path $RequestDir "summary.json"
    if (-not (Test-Path -LiteralPath $summaryPath)) {
        throw "missing request build summary: $summaryPath"
    }
    $code = @'
import json
import sys

data = json.load(open(sys.argv[2], encoding="utf-8-sig"))
index = int(sys.argv[3])
for item in data:
    request = item.get("request") or ""
    if request.endswith(f"full-{index:02d}.request.json"):
        print(json.dumps({
            "request": request,
            "appendedMissingEquations": int(item.get("appended_missing_equations") or 0),
            "questions": int(item.get("questions") or 0),
            "pngJpegImages": int(item.get("png_jpeg_images") or 0),
        }))
        break
else:
    raise SystemExit(f"request summary missing full-{index:02d}.request.json")
'@
    $encodedCode = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($code))
    $json = & python -c "import base64,sys; exec(base64.b64decode(sys.argv[1]).decode('utf-8'))" $encodedCode $summaryPath $Index
    if ($LASTEXITCODE -ne 0) {
        throw "python failed with exit code $LASTEXITCODE"
    }
    return $json | ConvertFrom-Json
}

function Read-MetricSummary {
    param(
        [string]$Path
    )
    $code = @'
import json
import sys

d = json.load(open(sys.argv[2], encoding="utf-8-sig"))
out = {
    "expectedOle": int(d["source_meta"]["objects"]),
    "generatedOle": int(d["generated_meta"]["objects"]),
    "generatedWmf": int(d["generated_meta"]["media_wmf"]),
    "pairedObjects": int(d["paired_objects"]),
    "unpairedSourceObjects": int(d["unpaired_source_objects"]),
    "unpairedGeneratedObjects": int(d["unpaired_generated_objects"]),
    "missingGeneratedObjects": int(d["missing_generated_objects"]),
    "extraGeneratedObjects": int(d["extra_generated_objects"]),
    "nonWmfGenerated": int(d["non_wmf_generated"]),
    "ordinalLatexMismatches": int(d["ordinal_latex_mismatches"]),
    "targetObjects": int(d["target_metric_objects"]),
    "targetWmfWidthWithin1Pct": int(d["target_wmf_width_ratio"]["within_1pct"]),
    "targetWmfHeightWithin1Pct": int(d["target_wmf_height_ratio"]["within_1pct"]),
    "targetShapeWidthWithin1Pct": int(d["target_shape_width_ratio"]["within_1pct"]),
    "targetShapeHeightWithin1Pct": int(d["target_shape_height_ratio"]["within_1pct"]),
    "targetWmfWidthN": int(d["target_wmf_width_ratio"]["n"]),
    "targetWmfHeightN": int(d["target_wmf_height_ratio"]["n"]),
    "targetShapeWidthN": int(d["target_shape_width_ratio"]["n"]),
    "targetShapeHeightN": int(d["target_shape_height_ratio"]["n"]),
    "targetWmfWidthMaxErrorPct": float(d["target_wmf_width_ratio"]["max_abs_error_pct"]),
    "targetWmfHeightMaxErrorPct": float(d["target_wmf_height_ratio"]["max_abs_error_pct"]),
    "targetShapeWidthMaxErrorPct": float(d["target_shape_width_ratio"]["max_abs_error_pct"]),
    "targetShapeHeightMaxErrorPct": float(d["target_shape_height_ratio"]["max_abs_error_pct"]),
}
print(json.dumps(out))
'@
    $encodedCode = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($code))
    $json = & python -c "import base64,sys; exec(base64.b64decode(sys.argv[1]).decode('utf-8'))" $encodedCode $Path
    if ($LASTEXITCODE -ne 0) {
        throw "python failed with exit code $LASTEXITCODE"
    }
    return $json | ConvertFrom-Json
}

if (-not $SkipGenerate) {
    if (-not $RequestDir) {
        Invoke-Checked "powershell" @(
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            (Join-Path $repoRoot "scripts\run_xsc_style_hints.ps1"),
            "-Start",
            "$Start",
            "-End",
            "$End",
            "-DatasetDir",
            $DatasetDir,
            "-AnalysisDir",
            $AnalysisDir,
            "-DocxToLatexDir",
            $DocxToLatexDir
        )
        $requestArgs = @(
            (Join-Path $repoRoot "scripts\make_full_batch10_requests.py"),
            "--latex-root",
            $LatexRoot,
            "--out-dir",
            $requestOut,
            "--style-hints-root",
            $styleHintsOut,
            "--start",
            "$Start",
            "--end",
            "$End"
        )
        if ($DatasetManifest) {
            $requestArgs += @("--manifest", $DatasetManifest)
        }
        Invoke-Checked "python" $requestArgs
    }
    Push-Location $repoRoot
    try {
        Invoke-Checked $mvn @(
            "-q",
            "-Dtest=XscFullBatch10DocxTest",
            "-Dxsc.analysis.dir=$AnalysisDir",
            "-Dxsc.full.request.dir=$requestOut",
            "-Dxsc.full.run.output.dir=$docxOut",
            "-Dxsc.full.start=$Start",
            "-Dxsc.full.end=$End",
            "-Dxsc.full.resume=true",
            "test"
        )
    } finally {
        Pop-Location
    }
}

if ($GenerateOnly) {
    $generationScanDir = Join-Path $strictOut "generation-only"
    New-Item -ItemType Directory -Force -Path $generationScanDir | Out-Null
    New-Item -ItemType Directory -Force -Path $runRoot | Out-Null
    $expectedTotal = 0
    $generatedTotal = 0
    $validTotal = 0
    $leakTotal = 0
    $documents = @()
    for ($i = $Start; $i -le $End; $i++) {
        $pad = "{0:D2}" -f $i
        $generatedDocx = Find-GeneratedDocx $docxOut $pad
        $reportPath = Join-Path $LatexRoot "$i\$i.report.json"
        if (-not (Test-Path -LiteralPath $generatedDocx)) { throw "missing generated docx: $generatedDocx" }
        $expected = $null
        if (Test-Path -LiteralPath $reportPath) {
            $sourceReport = Get-Content -LiteralPath $reportPath -Raw -Encoding UTF8 | ConvertFrom-Json
            $expected = @($sourceReport.equations).Count
        }
        $scanPath = Join-Path $generationScanDir "generated-$pad.json"
        $reuseScan = (Test-Path -LiteralPath $scanPath) -and
            ((Get-Item -LiteralPath $scanPath).LastWriteTimeUtc -ge (Get-Item -LiteralPath $generatedDocx).LastWriteTimeUtc)
        if (-not $reuseScan) {
            $scanArgs = @((Join-Path $repoRoot "scripts\scan_docx_latex_leaks.py"), $generatedDocx, "--out", $scanPath)
            if ($null -ne $expected) {
                $scanArgs += @("--expect-ole-count", "$expected")
            }
            & python @scanArgs | Out-Null
            if ($LASTEXITCODE -ne 0) { throw "generated OLE validation failed for doc $i; inspect $scanPath" }
        }
        $scan = Get-Content -LiteralPath $scanPath -Raw -Encoding UTF8 | ConvertFrom-Json
        if ([int]$scan.leakCount -ne 0 -or [int]$scan.invalidMathTypeOleCount -ne 0) {
            throw "generated OLE validation failed for doc $i; inspect $scanPath"
        }
        if ($null -ne $expected -and
                ([int]$scan.oleEmbeddingCount -ne [int]$expected -or
                 [int]$scan.validMathTypeOleCount -ne [int]$expected)) {
            throw "generated OLE count mismatch for doc $i; inspect $scanPath"
        }
        $observedExpected = if ($null -ne $expected) { [int]$expected } else { [int]$scan.oleEmbeddingCount }
        $expectedTotal += $observedExpected
        $generatedTotal += [int]$scan.oleEmbeddingCount
        $validTotal += [int]$scan.validMathTypeOleCount
        $leakTotal += [int]$scan.leakCount
        $documents += [pscustomobject]@{
            index = $i
            generated = $generatedDocx
            expectedOle = $observedExpected
            sourceReportBacked = $null -ne $expected
            generatedOle = [int]$scan.oleEmbeddingCount
            validOle = [int]$scan.validMathTypeOleCount
            latexLeaks = [int]$scan.leakCount
            scan = $scanPath
        }
    }
    $generationReport = [ordered]@{
        schemaVersion = 1
        mode = "project-generation-only"
        range = [ordered]@{ start = $Start; end = $End }
        expectedOleCount = $ExpectedOleTotal
        observedRequestOleCount = $expectedTotal
        generatedOleCount = $generatedTotal
        validOleCount = $validTotal
        latexLeakCount = $leakTotal
        pairTargetReached = $generatedTotal -ge $ExpectedOleTotal
        passed = $generatedTotal -gt 0 -and $generatedTotal -eq $validTotal -and $leakTotal -eq 0
        documents = $documents
    }
    $generationReportPath = Join-Path $runRoot "generation-only-summary.json"
    $generationReport | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $generationReportPath -Encoding UTF8
    Write-Host "generationOnlyReport=$generationReportPath"
    Write-Host "expectedOle=$ExpectedOleTotal observedRequestOle=$expectedTotal generatedOle=$generatedTotal validOle=$validTotal leaks=$leakTotal"
    if (-not $generationReport.passed) { throw "xsc generation-only gate failed; inspect $generationReportPath" }
    exit 0
}

New-Item -ItemType Directory -Force -Path $sizeOut | Out-Null
New-Item -ItemType Directory -Force -Path $strictOut | Out-Null
$formattedOut = Join-Path $runRoot "mathtype-formatted"
New-Item -ItemType Directory -Force -Path $formattedOut | Out-Null
$formatScript = Join-Path $repoRoot "scripts\format_word_mathtype_equations.ps1"
$strictItems = @()
for ($i = $Start; $i -le $End; $i++) {
    $pad = "{0:D2}" -f $i
    $sourceDocx = Resolve-SourceDocx $i
    $generatedDocx = Find-GeneratedDocx $docxOut $pad
    $reportPath = Join-Path $LatexRoot "$i\$i.report.json"
    if (-not (Test-Path -LiteralPath $reportPath)) {
        $fallbackReport = Join-Path $AnalysisDir "batch10-latex\$i\$i.report.json"
        if (Test-Path -LiteralPath $fallbackReport) {
            $reportPath = $fallbackReport
        }
    }
    if (-not (Test-Path -LiteralPath $sourceDocx)) { throw "missing source docx: $sourceDocx" }
    if (-not (Test-Path -LiteralPath $generatedDocx)) { throw "missing generated docx: $generatedDocx" }
    if (-not (Test-Path -LiteralPath $reportPath)) { throw "missing source report json: $reportPath" }
    $requestPath = Join-Path $requestOut "full-$pad.request.json"
    if (-not (Test-Path -LiteralPath $requestPath)) { throw "missing request json: $requestPath" }
    $formattedDocx = Join-Path $formattedOut ("generated-{0}.docx" -f $pad)
    $formatReport = Join-Path $formattedOut ("generated-{0}.format-report.json" -f $pad)
    $documentPreferencePath = if ($formatPreferencePath) { $formatPreferencePath } else { $sourceDocx }
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $formatScript `
        -InputDocx $generatedDocx -OutputDocx $formattedDocx -OutReport $formatReport `
        -PreferenceSource Clipboard -PreferenceDocx $documentPreferencePath | Out-Null
    if ($LASTEXITCODE -ne 0) { throw "MathType Format Equations failed for doc $i" }
    $geometryReport = Join-Path $formattedOut ("generated-{0}.geometry-report.json" -f $pad)
    # xsc is a DOCX round trip: keep MathType's rewritten OLE/WMF drawing records,
    # but preserve the source object's Word box and baseline exactly.
    $geometryResult = & python $restoreGeometryScript $sourceDocx $formattedDocx $formattedDocx
    if ($LASTEXITCODE -ne 0) { throw "formatted geometry restore failed for doc $i" }
    $geometryResult | Set-Content -LiteralPath $geometryReport -Encoding UTF8
    $generatedDocx = $formattedDocx
    $docSizeOut = Join-Path $sizeOut $pad
    New-Item -ItemType Directory -Force -Path $docSizeOut | Out-Null
    $compareArgs = @(
        $compareMetricsScript,
        $sourceDocx,
        $generatedDocx,
        $docSizeOut
    )
    $compareArgs += @("--source-report", $reportPath)
    $compareArgs += @("--generated-request", $requestPath)
    Invoke-Checked "python" $compareArgs -Quiet

    $summaryPath = Join-Path $docSizeOut "$([IO.Path]::GetFileNameWithoutExtension($sourceDocx))_summary.json"
    if (-not (Test-Path -LiteralPath $summaryPath)) {
        throw "missing pair metric summary: $summaryPath"
    }
    $requestBuildItem = Read-RequestBuildItem $requestOut $i
    $appendedMissing = [int]$requestBuildItem.appendedMissingEquations
    if ($appendedMissing -ne 0) {
        throw "request appended missing formulas for doc $i`: appendedMissingEquations=$appendedMissing request=$requestPath"
    }
    $metric = Read-MetricSummary $summaryPath
    $expectedOle = [int]$metric.expectedOle
    $generatedOle = [int]$metric.generatedOle
    $generatedWmf = [int]$metric.generatedWmf
    $pairedObjects = [int]$metric.pairedObjects
    $unpairedSource = [int]$metric.unpairedSourceObjects
    $unpairedGenerated = [int]$metric.unpairedGeneratedObjects
    $missingGenerated = [int]$metric.missingGeneratedObjects
    $extraGenerated = [int]$metric.extraGeneratedObjects
    $nonWmfGenerated = [int]$metric.nonWmfGenerated
    $ordinalLatexMismatches = [int]$metric.ordinalLatexMismatches
    if ($generatedOle -ne $expectedOle -or $generatedWmf -ne $expectedOle -or
        $pairedObjects -ne $expectedOle -or
        $unpairedSource -ne 0 -or $unpairedGenerated -ne 0 -or
        $missingGenerated -ne 0 -or $extraGenerated -ne 0 -or
        $nonWmfGenerated -ne 0 -or $ordinalLatexMismatches -ne 0) {
        throw "object/order gate failed for doc $i`: expected=$expectedOle generatedOle=$generatedOle generatedWmf=$generatedWmf paired=$pairedObjects unpaired=($unpairedSource,$unpairedGenerated) missingExtra=($missingGenerated,$extraGenerated) nonWmf=$nonWmfGenerated ordinalMismatches=$ordinalLatexMismatches"
    }
    $scanPath = Join-Path $strictOut "latex-leak-scan-$pad.json"
    Invoke-Checked "python" @(
        (Join-Path $repoRoot "scripts\scan_docx_latex_leaks.py"),
        $generatedDocx,
        "--expect-ole-count",
        "$expectedOle",
        "--expect-wmf-count",
        "$expectedOle",
        "--out",
        $scanPath
    ) -Quiet
    $targetObjects = [int]$metric.targetObjects
    $targetWmfW = [int]$metric.targetWmfWidthWithin1Pct
    $targetWmfH = [int]$metric.targetWmfHeightWithin1Pct
    $targetShapeW = [int]$metric.targetShapeWidthWithin1Pct
    $targetShapeH = [int]$metric.targetShapeHeightWithin1Pct
    $targetWmfWN = [int]$metric.targetWmfWidthN
    $targetWmfHN = [int]$metric.targetWmfHeightN
    $targetShapeWN = [int]$metric.targetShapeWidthN
    $targetShapeHN = [int]$metric.targetShapeHeightN
    $maxTargetWmfW = [double]$metric.targetWmfWidthMaxErrorPct
    $maxTargetWmfH = [double]$metric.targetWmfHeightMaxErrorPct
    $maxTargetShapeW = [double]$metric.targetShapeWidthMaxErrorPct
    $maxTargetShapeH = [double]$metric.targetShapeHeightMaxErrorPct
    if ($targetObjects -ne $expectedOle -or
        $targetWmfWN -ne $expectedOle -or $targetWmfHN -ne $expectedOle -or
        $targetShapeWN -ne $expectedOle -or $targetShapeHN -ne $expectedOle) {
        throw "incomplete target metrics for doc $i`: expected=$expectedOle targetObjects=$targetObjects wmfN=($targetWmfWN,$targetWmfHN) shapeN=($targetShapeWN,$targetShapeHN) request=$requestPath"
    }
    $strictItems += [pscustomobject]@{
        doc = $i
        source = $sourceDocx
        generated = $generatedDocx
        request = $requestPath
        expectedOle = $expectedOle
        generatedOle = $generatedOle
        generatedWmf = $generatedWmf
        pairedObjects = $pairedObjects
        unpairedSourceObjects = $unpairedSource
        unpairedGeneratedObjects = $unpairedGenerated
        missingGeneratedObjects = $missingGenerated
        extraGeneratedObjects = $extraGenerated
        nonWmfGenerated = $nonWmfGenerated
        ordinalLatexMismatches = $ordinalLatexMismatches
        appendedMissingEquations = $appendedMissing
        targetObjects = $targetObjects
        targetWmfWidthWithin1Pct = $targetWmfW
        targetWmfHeightWithin1Pct = $targetWmfH
        targetShapeWidthWithin1Pct = $targetShapeW
        targetShapeHeightWithin1Pct = $targetShapeH
        targetWmfWidthMaxErrorPct = $maxTargetWmfW
        targetWmfHeightMaxErrorPct = $maxTargetWmfH
        targetShapeWidthMaxErrorPct = $maxTargetShapeW
        targetShapeHeightMaxErrorPct = $maxTargetShapeH
    }
    if ($targetObjects -ne $targetWmfW -or $targetObjects -ne $targetWmfH -or
        $targetObjects -ne $targetShapeW -or $targetObjects -ne $targetShapeH) {
        throw "target physical size gate failed for doc $i`: targetObjects=$targetObjects wmf=($targetWmfW,$targetWmfH) shape=($targetShapeW,$targetShapeH)"
    }
    if ($maxTargetWmfW -gt $MaxTargetSizeErrorPct -or $maxTargetWmfH -gt $MaxTargetSizeErrorPct -or
        $maxTargetShapeW -gt $MaxTargetSizeErrorPct -or $maxTargetShapeH -gt $MaxTargetSizeErrorPct) {
        throw "target max error gate failed for doc $i`: wmf=($maxTargetWmfW,$maxTargetWmfH) shape=($maxTargetShapeW,$maxTargetShapeH) limit=$MaxTargetSizeErrorPct"
    }
}

$allTargetPhysicalSizesWithinLimit = @($strictItems | Where-Object {
    $_.targetObjects -ne $_.targetWmfWidthWithin1Pct -or
    $_.targetObjects -ne $_.targetWmfHeightWithin1Pct -or
    $_.targetObjects -ne $_.targetShapeWidthWithin1Pct -or
    $_.targetObjects -ne $_.targetShapeHeightWithin1Pct
}).Count -eq 0
$strictSummary = [pscustomobject]@{
    stamp = $Stamp
    range = @{ start = $Start; end = $End }
    maxTargetSizeErrorPct = $MaxTargetSizeErrorPct
    allTargetPhysicalSizesWithinLimit = $allTargetPhysicalSizesWithinLimit
    items = $strictItems
}
$strictSummaryPath = Join-Path $strictOut "strict-acceptance-summary.json"
$strictSummary | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $strictSummaryPath -Encoding UTF8
Write-Host "strictAcceptance=$strictSummaryPath"

if ($SkipMtef) {
    Write-Host "skip mtef comparison"
    exit 0
}

New-Item -ItemType Directory -Force -Path $mtefOut | Out-Null
New-Item -ItemType Directory -Force -Path $exactMtefOut | Out-Null
if (-not (Test-Path -LiteralPath $DocxToLatexDir)) {
    throw "missing docxtolatex dir: $DocxToLatexDir"
}
$classpathFile = Join-Path $exactMtefOut "runtime-classpath.txt"
Push-Location $repoRoot
try {
    Invoke-Checked $mvn @(
        "-q",
        "-DskipTests",
        "compile",
        "dependency:build-classpath",
        "-Dmdep.outputFile=$classpathFile"
    ) -Quiet
} finally {
    Pop-Location
}
$runtimeClasspath = (Join-Path $repoRoot "target\classes") + [IO.Path]::PathSeparator +
    (Get-Content -LiteralPath $classpathFile -Raw -Encoding UTF8).Trim()
$exactItems = @()
Push-Location $DocxToLatexDir
try {
    for ($i = $Start; $i -le $End; $i++) {
        $pad = "{0:D2}" -f $i
        $sourceDocx = Resolve-SourceDocx $i
        $sourceStem = [IO.Path]::GetFileNameWithoutExtension($sourceDocx)
        $reportPath = Join-Path $LatexRoot "$i\$i.report.json"
        if (-not (Test-Path -LiteralPath $reportPath)) {
            $fallbackReport = Join-Path $AnalysisDir "batch10-latex\$i\$i.report.json"
            if (Test-Path -LiteralPath $fallbackReport) {
                $reportPath = $fallbackReport
            }
        }
        if (-not (Test-Path -LiteralPath $reportPath)) { throw "missing report json: $reportPath" }
        $generatedDocx = Find-GeneratedDocx $docxOut $pad
        $exactReportPath = Join-Path $exactMtefOut "$pad-exact-mtef.json"
        Invoke-Checked "java" @(
            "-cp",
            $runtimeClasspath,
            "com.lz.paperword.core.mtef.MathTypeDocxComparatorCli",
            $sourceDocx,
            $generatedDocx,
            $exactReportPath
        ) -Quiet
        $exactReport = Get-Content -LiteralPath $exactReportPath -Raw -Encoding UTF8 | ConvertFrom-Json
        $exactItems += [pscustomobject]@{
            doc = $i
            expectedObjects = [int]$exactReport.standardObjectCount
            generatedObjects = [int]$exactReport.generatedObjectCount
            validGeneratedOle = [int]$exactReport.validGeneratedOleCount
            normalizedStructureEqual = [int]$exactReport.normalizedStructureEqualCount
            rawMtefEqual = [int]$exactReport.rawMtefEqualCount
            report = $exactReportPath
        }
        Invoke-Checked "go" @(
            "run",
            $compareMtefScript,
            $sourceDocx,
            $generatedDocx,
            (Join-Path $mtefOut "$i"),
            $reportPath,
            (Join-Path $requestOut "full-$pad.request.json")
        ) -Quiet
        $mtefSummaryPath = Join-Path (Join-Path $mtefOut "$i") "$sourceStem`_mtef_summary.json"
        if (-not (Test-Path -LiteralPath $mtefSummaryPath)) { throw "missing mtef summary: $mtefSummaryPath" }
        $mtefCounts = Invoke-Checked "python" @(
            "-c",
            "import json,sys; d=json.load(open(sys.argv[1],encoding='utf-8')); print(str(d.get('pairedObjects',0))+' '+str(d.get('sourceObjects',0)))",
            $mtefSummaryPath
        )
        $mtefCountParts = @($mtefCounts -split "\s+")
        if ([int]$mtefCountParts[0] -ne [int]$mtefCountParts[1]) {
            throw "incomplete mtef pairs for doc $i`: paired=$($mtefCountParts[0]) source=$($mtefCountParts[1])"
        }
    }
} finally {
    Pop-Location
}

$exactExpected = ($exactItems | Measure-Object -Property expectedObjects -Sum).Sum
$exactGenerated = ($exactItems | Measure-Object -Property generatedObjects -Sum).Sum
$exactValid = ($exactItems | Measure-Object -Property validGeneratedOle -Sum).Sum
$exactEqual = ($exactItems | Measure-Object -Property normalizedStructureEqual -Sum).Sum
$exactRawEqual = ($exactItems | Measure-Object -Property rawMtefEqual -Sum).Sum
$exactSummary = [ordered]@{
    expectedObjects = [int]$exactExpected
    generatedObjects = [int]$exactGenerated
    validGeneratedOle = [int]$exactValid
    normalizedStructureEqual = [int]$exactEqual
    rawMtefEqual = [int]$exactRawEqual
    passed = $exactExpected -eq $exactGenerated -and $exactGenerated -eq $exactValid -and $exactGenerated -eq $exactEqual
    items = $exactItems
}
$exactSummaryPath = Join-Path $exactMtefOut "exact-mtef-summary.json"
$exactSummary | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $exactSummaryPath -Encoding UTF8

Invoke-Checked "python" @(
    (Join-Path $repoRoot "scripts\summarize_xsc_acceptance.py"),
    $Stamp,
    $sizeOut,
    $mtefOut,
    "--start",
    "$Start",
    "--end",
    "$End",
    "--analysis-dir",
    $AnalysisDir,
    "--exact-mtef-summary",
    $exactSummaryPath
)
if (-not $exactSummary.passed) {
    throw "exact normalized MTEF gate failed: expected=$exactExpected generated=$exactGenerated validOle=$exactValid exact=$exactEqual report=$exactSummaryPath"
}
