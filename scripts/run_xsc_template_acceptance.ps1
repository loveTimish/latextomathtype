param(
    [int]$Start = 1,
    [int]$End = 10,
    [string]$Stamp = "",
    [string]$DatasetDir = "",
    [string]$AnalysisDir = "J:\latextomathtype\analysis",
    [string]$LatexRoot = "J:\latextomathtype\analysis\xsc-latex-fixed",
    [string]$DocxToLatexDir = "J:\docxtolatex\docxtolatex-main",
    [string]$GeneratedDir = "",
    [string]$RequestDir = "",
    [string]$DatasetManifest = "",
    [double]$MaxSizeErrorPct = 1.0,
    [switch]$SkipGenerate
)

$ErrorActionPreference = "Stop"

if (-not $DatasetDir) {
    $DatasetDir = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String("Rjpc6LWE5paZXHhzY+i1hOaWmVx3b3JkX2ZpbGVz"))
}

if (-not $Stamp) {
    $Stamp = Get-Date -Format "yyyyMMdd-HHmmss"
}

$repoRoot = Resolve-Path (Join-Path $PSScriptRoot "..")
$mvn = Join-Path $repoRoot ".mvn\apache-maven-3.9.12\bin\mvn.cmd"
if (-not (Test-Path -LiteralPath $mvn)) {
    throw "Bundled Maven not found: $mvn"
}
$generatedDir = if ($GeneratedDir) { $GeneratedDir } else { Join-Path $AnalysisDir "batch10-full-docx\$Stamp" }
$templateDir = Join-Path $AnalysisDir "template-rebuild\$Stamp"
$pairDir = Join-Path $templateDir "mtef-pairs"
$recordDir = Join-Path $templateDir "wmf-records"
$metricDir = Join-Path $templateDir "metrics"
$generatedMetricDir = Join-Path $templateDir "generated-metrics"
$requestOut = if ($RequestDir) { $RequestDir } else { Join-Path $AnalysisDir "batch10-full-requests" }
$compareMetricsScript = Join-Path $repoRoot "scripts\compare_docx_pair_metrics.py"
if (-not (Test-Path -LiteralPath $compareMetricsScript)) {
    throw "missing compare metrics script: $compareMetricsScript"
}
$compareMtefScript = Join-Path $repoRoot "scripts\compare_mtef_native.go"
if (-not (Test-Path -LiteralPath $compareMtefScript)) {
    throw "missing compare mtef script: $compareMtefScript"
}
if (-not (Test-Path -LiteralPath $DocxToLatexDir)) {
    throw "missing docxtolatex dir: $DocxToLatexDir"
}
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
    $code = @'
import json
import sys

data = json.load(open(sys.argv[1], encoding="utf-8-sig"))
if isinstance(data, dict):
    data = [data]
for item in data:
    index = item.get("index")
    source_path = item.get("sourcePath") or ""
    source_name = item.get("sourceName") or ""
    if index is not None and source_path:
        print(f"{int(index)}\t{source_path}\t{source_name}")
'@
    $lines = & python -c $code $Path
    if ($LASTEXITCODE -ne 0) {
        throw "python failed with exit code $LASTEXITCODE"
    }
    return @($lines | ForEach-Object {
        $parts = $_ -split "`t", 3
        [pscustomobject]@{
            index = [int]$parts[0]
            sourcePath = $parts[1]
            sourceName = if ($parts.Count -gt 2) { $parts[2] } else { "" }
        }
    })
}

function Read-MetricSummary {
    param(
        [string]$Path
    )
    if (-not (Test-Path -LiteralPath $Path)) {
        throw "missing metric summary: $Path"
    }
    $code = @'
import json
import sys

d = json.load(open(sys.argv[2], encoding="utf-8-sig"))
def ratio_count(name, field, default=0):
    return int((d.get(name) or {}).get(field, default))

def ratio_float(name, field, default=0.0):
    return float((d.get(name) or {}).get(field, default))

out = {
    "sourceObjects": int(d["source_meta"]["objects"]),
    "sourceEmbeddings": int(d["source_meta"].get("embeddings_bin") or 0),
    "generatedObjects": int(d["generated_meta"]["objects"]),
    "generatedEmbeddings": int(d["generated_meta"].get("embeddings_bin") or 0),
    "generatedWmf": int(d["generated_meta"]["media_wmf"]),
    "pairedObjects": int(d["paired_objects"]),
    "unpairedSourceObjects": int(d["unpaired_source_objects"]),
    "unpairedGeneratedObjects": int(d["unpaired_generated_objects"]),
    "missingGeneratedObjects": int(d["missing_generated_objects"]),
    "extraGeneratedObjects": int(d["extra_generated_objects"]),
    "nonWmfGenerated": int(d["non_wmf_generated"]),
    "wmfWidthWithin1pct": int(d["wmf_width_ratio"]["within_1pct"]),
    "wmfHeightWithin1pct": int(d["wmf_height_ratio"]["within_1pct"]),
    "wmfWidthMaxErrorPct": float(d["wmf_width_ratio"]["max_abs_error_pct"]),
    "wmfHeightMaxErrorPct": float(d["wmf_height_ratio"]["max_abs_error_pct"]),
    "shapeWidthMaxErrorPct": float(d["shape_width_ratio"]["max_abs_error_pct"]),
    "shapeHeightMaxErrorPct": float(d["shape_height_ratio"]["max_abs_error_pct"]),
    "shapeWidthN": ratio_count("shape_width_ratio", "n"),
    "shapeHeightN": ratio_count("shape_height_ratio", "n"),
    "shapeWidthWithin1pct": ratio_count("shape_width_ratio", "within_1pct"),
    "shapeHeightWithin1pct": ratio_count("shape_height_ratio", "within_1pct"),
    "targetObjects": int(d.get("target_metric_objects") or 0),
    "targetWmfWidthN": ratio_count("target_wmf_width_ratio", "n"),
    "targetWmfHeightN": ratio_count("target_wmf_height_ratio", "n"),
    "targetShapeWidthN": ratio_count("target_shape_width_ratio", "n"),
    "targetShapeHeightN": ratio_count("target_shape_height_ratio", "n"),
    "targetWmfWidthWithin1pct": ratio_count("target_wmf_width_ratio", "within_1pct"),
    "targetWmfHeightWithin1pct": ratio_count("target_wmf_height_ratio", "within_1pct"),
    "targetShapeWidthWithin1pct": ratio_count("target_shape_width_ratio", "within_1pct"),
    "targetShapeHeightWithin1pct": ratio_count("target_shape_height_ratio", "within_1pct"),
    "targetWmfWidthMaxErrorPct": ratio_float("target_wmf_width_ratio", "max_abs_error_pct"),
    "targetWmfHeightMaxErrorPct": ratio_float("target_wmf_height_ratio", "max_abs_error_pct"),
    "targetShapeWidthMaxErrorPct": ratio_float("target_shape_width_ratio", "max_abs_error_pct"),
    "targetShapeHeightMaxErrorPct": ratio_float("target_shape_height_ratio", "max_abs_error_pct"),
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

$manifestItems = @(Read-ManifestItems $DatasetManifest)

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

if (-not $SkipGenerate) {
    Push-Location $repoRoot
    try {
        Invoke-Checked $mvn @(
            "-q",
            "-Dtest=XscFullBatch10DocxTest#generateTenFullXscDocxFiles",
            "-Dxsc.analysis.dir=$AnalysisDir",
            "-Dxsc.full.request.dir=$requestOut",
            "-Dxsc.full.run.output.dir=$generatedDir",
            "-Dxsc.full.start=$Start",
            "-Dxsc.full.end=$End",
            "test"
        )
    } finally {
        Pop-Location
    }
}

New-Item -ItemType Directory -Force -Path $templateDir, $pairDir, $recordDir, $metricDir, $generatedMetricDir | Out-Null

$items = @()
Push-Location $DocxToLatexDir
try {
    for ($i = $Start; $i -le $End; $i++) {
        $pad = "{0:D2}" -f $i
        $sourceDocx = Resolve-SourceDocx $i
        $generatedMatches = @(Get-ChildItem -LiteralPath $generatedDir -Filter "*_$pad.docx" -ErrorAction SilentlyContinue)
        if ($generatedMatches.Count -eq 1) {
            $generatedDocx = $generatedMatches[0].FullName
        } elseif ($generatedMatches.Count -gt 1) {
            throw "ambiguous generated docx for $pad in $generatedDir`: $($generatedMatches.Name -join ', ')"
        } else {
            $generatedDocx = Join-Path $generatedDir "$pad.docx"
        }
        $reportPath = Join-Path $LatexRoot "$i\$i.report.json"
        if (-not (Test-Path -LiteralPath $reportPath)) {
            $fallbackReport = Join-Path $AnalysisDir "batch10-latex\$i\$i.report.json"
            if (Test-Path -LiteralPath $fallbackReport) {
                $reportPath = $fallbackReport
            }
        }
        $requestPath = Join-Path $requestOut "full-$pad.request.json"
        $pairOut = Join-Path $pairDir "$i"
        $templateDocx = Join-Path $templateDir "xsc-template-rebuild_$pad.docx"
        $recordJson = Join-Path $recordDir "$i-wmf-records.json"
        $metricOut = Join-Path $metricDir "$i"
        $generatedMetricOut = Join-Path $generatedMetricDir "$i"
        New-Item -ItemType Directory -Force -Path $pairOut, $metricOut, $generatedMetricOut | Out-Null

        if (-not (Test-Path -LiteralPath $sourceDocx)) { throw "missing source docx: $sourceDocx" }
        if (-not (Test-Path -LiteralPath $generatedDocx)) { throw "missing generated docx: $generatedDocx" }
        if (-not (Test-Path -LiteralPath $reportPath)) { throw "missing report json: $reportPath" }
        if (-not (Test-Path -LiteralPath $requestPath)) { throw "missing request json: $requestPath" }

        Invoke-Checked "go" @(
            "run",
            $compareMtefScript,
            $sourceDocx,
            $generatedDocx,
            $pairOut,
            $reportPath,
            $requestPath
        ) -Quiet
        $sourceStem = [IO.Path]::GetFileNameWithoutExtension($sourceDocx)
        $pairCsv = Join-Path $pairOut "$sourceStem`_mtef_pairs.csv"
        if (-not (Test-Path -LiteralPath $pairCsv)) { throw "missing pair csv: $pairCsv" }

        Invoke-Checked "python" @(
            $compareMetricsScript,
            $sourceDocx,
            $generatedDocx,
            $generatedMetricOut,
            "--generated-request",
            $requestPath
        ) -Quiet
        $generatedMetricSummary = Get-ChildItem -LiteralPath $generatedMetricOut -Filter "*_summary.json" | Select-Object -First 1
        if (-not $generatedMetricSummary) { throw "missing generated metric summary for doc $i" }
        $generatedMetric = Read-MetricSummary $generatedMetricSummary.FullName
        $generatedSourceObjects = [int]$generatedMetric.sourceObjects
        $generatedObjects = [int]$generatedMetric.generatedObjects
        $generatedEmbeddings = [int]$generatedMetric.generatedEmbeddings
        $generatedWmf = [int]$generatedMetric.generatedWmf
        $generatedPairedObjects = [int]$generatedMetric.pairedObjects
        $generatedUnpairedSource = [int]$generatedMetric.unpairedSourceObjects
        $generatedUnpairedGenerated = [int]$generatedMetric.unpairedGeneratedObjects
        $generatedMissing = [int]$generatedMetric.missingGeneratedObjects
        $generatedExtra = [int]$generatedMetric.extraGeneratedObjects
        $generatedNonWmf = [int]$generatedMetric.nonWmfGenerated
        $generatedTargetObjects = [int]$generatedMetric.targetObjects
        $generatedTargetWmfWidthN = [int]$generatedMetric.targetWmfWidthN
        $generatedTargetWmfHeightN = [int]$generatedMetric.targetWmfHeightN
        $generatedTargetShapeWidthN = [int]$generatedMetric.targetShapeWidthN
        $generatedTargetShapeHeightN = [int]$generatedMetric.targetShapeHeightN
        $generatedTargetWmfWidthWithin = [int]$generatedMetric.targetWmfWidthWithin1pct
        $generatedTargetWmfHeightWithin = [int]$generatedMetric.targetWmfHeightWithin1pct
        $generatedTargetShapeWidthWithin = [int]$generatedMetric.targetShapeWidthWithin1pct
        $generatedTargetShapeHeightWithin = [int]$generatedMetric.targetShapeHeightWithin1pct
        $generatedTargetWmfWidthMax = [double]$generatedMetric.targetWmfWidthMaxErrorPct
        $generatedTargetWmfHeightMax = [double]$generatedMetric.targetWmfHeightMaxErrorPct
        $generatedTargetShapeWidthMax = [double]$generatedMetric.targetShapeWidthMaxErrorPct
        $generatedTargetShapeHeightMax = [double]$generatedMetric.targetShapeHeightMaxErrorPct
        if ($generatedSourceObjects -ne $generatedObjects -or
            $generatedPairedObjects -ne $generatedSourceObjects -or
            $generatedUnpairedSource -ne 0 -or
            $generatedUnpairedGenerated -ne 0 -or
            $generatedMissing -ne 0 -or
            $generatedExtra -ne 0) {
            throw "template preflight failed for doc $i`: sourceObjects=$generatedSourceObjects generatedObjects=$generatedObjects paired=$generatedPairedObjects unpaired=($generatedUnpairedSource,$generatedUnpairedGenerated) missing=$generatedMissing extra=$generatedExtra"
        }
        if ($generatedEmbeddings -ne $generatedObjects) {
            throw "template preflight failed for doc $i`: generatedEmbeddings=$generatedEmbeddings generatedObjects=$generatedObjects"
        }
        if ($generatedWmf -ne $generatedObjects -or $generatedNonWmf -ne 0) {
            throw "template preflight failed for doc $i`: generatedWmf=$generatedWmf generatedObjects=$generatedObjects nonWmfGenerated=$generatedNonWmf"
        }
        if ($generatedTargetObjects -ne $generatedObjects -or
            $generatedTargetWmfWidthN -ne $generatedObjects -or
            $generatedTargetWmfHeightN -ne $generatedObjects -or
            $generatedTargetShapeWidthN -ne $generatedObjects -or
            $generatedTargetShapeHeightN -ne $generatedObjects) {
            throw "template preflight failed for doc $i`: incomplete target metrics targetObjects=$generatedTargetObjects generatedObjects=$generatedObjects wmfN=($generatedTargetWmfWidthN,$generatedTargetWmfHeightN) shapeN=($generatedTargetShapeWidthN,$generatedTargetShapeHeightN)"
        }
        if ($generatedTargetWmfWidthWithin -ne $generatedObjects -or
            $generatedTargetWmfHeightWithin -ne $generatedObjects -or
            $generatedTargetShapeWidthWithin -ne $generatedObjects -or
            $generatedTargetShapeHeightWithin -ne $generatedObjects) {
            throw "template preflight failed for doc $i`: target physical size within1pct=($generatedTargetWmfWidthWithin,$generatedTargetWmfHeightWithin,$generatedTargetShapeWidthWithin,$generatedTargetShapeHeightWithin) generatedObjects=$generatedObjects"
        }
        if ($generatedTargetWmfWidthMax -gt $MaxSizeErrorPct -or
            $generatedTargetWmfHeightMax -gt $MaxSizeErrorPct -or
            $generatedTargetShapeWidthMax -gt $MaxSizeErrorPct -or
            $generatedTargetShapeHeightMax -gt $MaxSizeErrorPct) {
            throw "template preflight failed for doc $i`: target maxError wmf=($generatedTargetWmfWidthMax,$generatedTargetWmfHeightMax) shape=($generatedTargetShapeWidthMax,$generatedTargetShapeHeightMax) limit=$MaxSizeErrorPct"
        }

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
            $compareMetricsScript,
            $sourceDocx,
            $templateDocx,
            $metricOut
        ) -Quiet

        $record = Get-Content -LiteralPath $recordJson -Raw | ConvertFrom-Json
        $metricSummary = Get-ChildItem -LiteralPath $metricOut -Filter "*_summary.json" | Select-Object -First 1
        if (-not $metricSummary) { throw "missing template metric summary for doc $i" }
        $metric = Read-MetricSummary $metricSummary.FullName
        $items += [pscustomobject]@{
            doc = $i
            source = $sourceDocx
            generated = $generatedDocx
            template = $templateDocx
            pairCsv = $pairCsv
            wmf = [int]$record.wmf
            vectorText = [int]$record.vectorText
            stretchDib = [int]$record.stretchDib
            sourceObjects = [int]$metric.sourceObjects
            generatedObjects = [int]$metric.generatedObjects
            pairedObjects = [int]$metric.pairedObjects
            unpairedSourceObjects = [int]$metric.unpairedSourceObjects
            unpairedGeneratedObjects = [int]$metric.unpairedGeneratedObjects
            missingGeneratedObjects = [int]$metric.missingGeneratedObjects
            extraGeneratedObjects = [int]$metric.extraGeneratedObjects
            nonWmfGenerated = [int]$metric.nonWmfGenerated
            wmfWidthWithin1pct = [int]$metric.wmfWidthWithin1pct
            wmfHeightWithin1pct = [int]$metric.wmfHeightWithin1pct
            shapeWidthN = [int]$metric.shapeWidthN
            shapeHeightN = [int]$metric.shapeHeightN
            shapeWidthWithin1pct = [int]$metric.shapeWidthWithin1pct
            shapeHeightWithin1pct = [int]$metric.shapeHeightWithin1pct
            wmfWidthMaxErrorPct = [double]$metric.wmfWidthMaxErrorPct
            wmfHeightMaxErrorPct = [double]$metric.wmfHeightMaxErrorPct
            shapeWidthMaxErrorPct = [double]$metric.shapeWidthMaxErrorPct
            shapeHeightMaxErrorPct = [double]$metric.shapeHeightMaxErrorPct
        }
        $item = $items[-1]
        if ($item.pairedObjects -le 0) {
            throw "template acceptance failed for doc $i`: no paired formula objects"
        }
        if ($item.sourceObjects -ne $item.generatedObjects -or
            $item.pairedObjects -ne $item.sourceObjects -or
            $item.unpairedSourceObjects -ne 0 -or
            $item.unpairedGeneratedObjects -ne 0 -or
            $item.missingGeneratedObjects -ne 0 -or
            $item.extraGeneratedObjects -ne 0) {
            throw "template acceptance failed for doc $i`: sourceObjects=$($item.sourceObjects) generatedObjects=$($item.generatedObjects) paired=$($item.pairedObjects) unpaired=($($item.unpairedSourceObjects),$($item.unpairedGeneratedObjects)) missing=$($item.missingGeneratedObjects) extra=$($item.extraGeneratedObjects)"
        }
        if ($item.shapeWidthN -le 0) {
            throw "template acceptance failed for doc $i`: no shape width metrics"
        }
        if ($item.shapeHeightN -ne $item.pairedObjects) {
            throw "template acceptance failed for doc $i`: shapeHeightN=$($item.shapeHeightN) pairedObjects=$($item.pairedObjects)"
        }
        if ($item.shapeWidthWithin1pct -ne $item.shapeWidthN) {
            throw "template acceptance failed for doc $i`: shapeWidthWithin1pct=$($item.shapeWidthWithin1pct) shapeWidthN=$($item.shapeWidthN)"
        }
        if ($item.shapeHeightWithin1pct -ne $item.shapeHeightN) {
            throw "template acceptance failed for doc $i`: shapeHeightWithin1pct=$($item.shapeHeightWithin1pct) shapeHeightN=$($item.shapeHeightN)"
        }
        if ($item.shapeWidthMaxErrorPct -gt $MaxSizeErrorPct -or $item.shapeHeightMaxErrorPct -gt $MaxSizeErrorPct) {
            throw "template acceptance failed for doc $i`: shape maxError=($($item.shapeWidthMaxErrorPct),$($item.shapeHeightMaxErrorPct)) limit=$MaxSizeErrorPct"
        }
    }
} finally {
    Pop-Location
}

$summary = [pscustomobject]@{
    stamp = $Stamp
    range = @{ start = $Start; end = $End }
    maxSizeErrorPct = $MaxSizeErrorPct
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
