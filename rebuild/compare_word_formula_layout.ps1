param(
    [string]$ReferenceDocx = "rebuild-assets\external\fraction-split-reference.docx",
    [string]$GeneratedDocx = "target\reference-roundtrip\fraction-split-reference-regenerated.docx",
    [string]$OutJson = "target\reference-roundtrip\word-formula-layout-comparison.json",
    [string]$OutText = "target\reference-roundtrip\word-formula-layout-comparison.txt"
)

$ErrorActionPreference = "Stop"

function Resolve-ProjectPath([string]$PathValue) {
    if ([IO.Path]::IsPathRooted($PathValue)) {
        return $PathValue
    }
    return Join-Path (Get-Location) $PathValue
}

function Get-Stat([double[]]$Values) {
    if (-not $Values -or $Values.Count -eq 0) {
        return [PSCustomObject]@{ n = 0; avg = $null; min = $null; max = $null }
    }
    $measure = $Values | Measure-Object -Average -Minimum -Maximum
    return [PSCustomObject]@{
        n = $Values.Count
        avg = [math]::Round($measure.Average, 2)
        min = [math]::Round($measure.Minimum, 2)
        max = [math]::Round($measure.Maximum, 2)
    }
}

function Get-MathTypeInlineLayout([string]$DocxPath) {
    $resolved = (Resolve-Path -LiteralPath (Resolve-ProjectPath $DocxPath)).Path
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    try {
        $doc = $word.Documents.Open($resolved, $false, $true)
        try {
            $items = @()
            for ($i = 1; $i -le $doc.InlineShapes.Count; $i++) {
                $shape = $doc.InlineShapes.Item($i)
                $progId = $null
                try {
                    $progId = $shape.OLEFormat.ProgID
                } catch {
                    $progId = $null
                }
                if ($progId -ne "Equation.DSMT4") {
                    continue
                }
                $range = $shape.Range
                $items += [PSCustomObject]@{
                    Index = $items.Count
                    InlineShapeIndex = $i
                    Page = [int]$range.Information(3)
                    XPt = [double]$range.Information(5)
                    YPt = [double]$range.Information(6)
                    WidthPt = [double]$shape.Width
                    HeightPt = [double]$shape.Height
                }
            }
            return $items
        } finally {
            $doc.Close($false)
        }
    } finally {
        $word.Quit()
    }
}

$reference = @(Get-MathTypeInlineLayout $ReferenceDocx)
$generated = @(Get-MathTypeInlineLayout $GeneratedDocx)
$paired = [math]::Min($reference.Count, $generated.Count)
$rows = @()
for ($i = 0; $i -lt $paired; $i++) {
    $ref = $reference[$i]
    $gen = $generated[$i]
    $rows += [PSCustomObject]@{
        Index = $i
        ReferencePage = $ref.Page
        GeneratedPage = $gen.Page
        PageDelta = $gen.Page - $ref.Page
        ReferenceXPt = [math]::Round($ref.XPt, 2)
        GeneratedXPt = [math]::Round($gen.XPt, 2)
        AbsXDeltaPt = [math]::Round([math]::Abs($gen.XPt - $ref.XPt), 2)
        ReferenceYPt = [math]::Round($ref.YPt, 2)
        GeneratedYPt = [math]::Round($gen.YPt, 2)
        AbsYDeltaPt = [math]::Round([math]::Abs($gen.YPt - $ref.YPt), 2)
        ReferenceWidthPt = [math]::Round($ref.WidthPt, 2)
        GeneratedWidthPt = [math]::Round($gen.WidthPt, 2)
        AbsWidthDeltaPt = [math]::Round([math]::Abs($gen.WidthPt - $ref.WidthPt), 2)
        ReferenceHeightPt = [math]::Round($ref.HeightPt, 2)
        GeneratedHeightPt = [math]::Round($gen.HeightPt, 2)
        AbsHeightDeltaPt = [math]::Round([math]::Abs($gen.HeightPt - $ref.HeightPt), 2)
    }
}

$pageMismatches = @($rows | Where-Object { $_.PageDelta -ne 0 })
$summary = [PSCustomObject]@{
    Reference = (Resolve-Path -LiteralPath (Resolve-ProjectPath $ReferenceDocx)).Path
    Generated = (Resolve-Path -LiteralPath (Resolve-ProjectPath $GeneratedDocx)).Path
    ReferenceCount = $reference.Count
    GeneratedCount = $generated.Count
    PairedCount = $paired
    PageMismatchCount = $pageMismatches.Count
    AbsXDeltaPt = Get-Stat @($rows | ForEach-Object { [double]$_.AbsXDeltaPt })
    AbsYDeltaPt = Get-Stat @($rows | ForEach-Object { [double]$_.AbsYDeltaPt })
    AbsWidthDeltaPt = Get-Stat @($rows | ForEach-Object { [double]$_.AbsWidthDeltaPt })
    AbsHeightDeltaPt = Get-Stat @($rows | ForEach-Object { [double]$_.AbsHeightDeltaPt })
}

$payload = [PSCustomObject]@{
    Summary = $summary
    WorstX = @($rows | Sort-Object AbsXDeltaPt -Descending | Select-Object -First 12)
    WorstY = @($rows | Sort-Object AbsYDeltaPt -Descending | Select-Object -First 12)
    PageMismatches = @($pageMismatches | Select-Object -First 40)
    Rows = $rows
}

$outJsonPath = Resolve-ProjectPath $OutJson
$outTextPath = Resolve-ProjectPath $OutText
New-Item -ItemType Directory -Force -Path ([IO.Path]::GetDirectoryName($outJsonPath)) | Out-Null
$payload | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $outJsonPath -Encoding UTF8

$lines = @()
$lines += "Word formula layout comparison"
$lines += ""
$lines += "reference: $($summary.Reference)"
$lines += "generated: $($summary.Generated)"
$lines += ""
$lines += "counts: reference=$($summary.ReferenceCount) generated=$($summary.GeneratedCount) paired=$($summary.PairedCount)"
$lines += "page_mismatch_count=$($summary.PageMismatchCount)"
$lines += "abs_x_delta_pt=$($summary.AbsXDeltaPt | ConvertTo-Json -Compress)"
$lines += "abs_y_delta_pt=$($summary.AbsYDeltaPt | ConvertTo-Json -Compress)"
$lines += "abs_width_delta_pt=$($summary.AbsWidthDeltaPt | ConvertTo-Json -Compress)"
$lines += "abs_height_delta_pt=$($summary.AbsHeightDeltaPt | ConvertTo-Json -Compress)"
$lines += ""
$lines += "Worst Y deltas"
foreach ($row in ($payload.WorstY | Select-Object -First 8)) {
    $lines += "  #$($row.Index): ref=page$($row.ReferencePage) y=$($row.ReferenceYPt) gen=page$($row.GeneratedPage) y=$($row.GeneratedYPt) dy=$($row.AbsYDeltaPt)pt"
}
$lines += ""
$lines += "Worst X deltas"
foreach ($row in ($payload.WorstX | Select-Object -First 8)) {
    $lines += "  #$($row.Index): ref=page$($row.ReferencePage) x=$($row.ReferenceXPt) gen=page$($row.GeneratedPage) x=$($row.GeneratedXPt) dx=$($row.AbsXDeltaPt)pt"
}
$lines | Set-Content -LiteralPath $outTextPath -Encoding UTF8

Get-Content -LiteralPath $outTextPath
