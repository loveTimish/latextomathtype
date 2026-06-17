param(
    [Parameter(Mandatory = $true)]
    [string]$OutDocx,
    [Parameter(Mandatory = $true)]
    [string[]]$Formula,
    [string]$OutSourceReport = "",
    [switch]$Visible
)

$ErrorActionPreference = "Stop"

$resolvedDocx = [IO.Path]::GetFullPath((Join-Path (Get-Location) $OutDocx))
$docxDir = Split-Path -Parent $resolvedDocx
New-Item -ItemType Directory -Force -Path $docxDir | Out-Null

if ([string]::IsNullOrWhiteSpace($OutSourceReport)) {
    $OutSourceReport = [IO.Path]::ChangeExtension($resolvedDocx, ".source-report.json")
}
$resolvedReport = [IO.Path]::GetFullPath((Join-Path (Get-Location) $OutSourceReport))
New-Item -ItemType Directory -Force -Path (Split-Path -Parent $resolvedReport) | Out-Null

$word = New-Object -ComObject Word.Application
$word.Visible = [bool]$Visible
$word.DisplayAlerts = 0
$word.ScreenUpdating = [bool]$Visible
$doc = $null
try {
    $doc = $word.Documents.Add()
    $equations = @()
    $mathTypeOleCount = 0
    foreach ($tex in $Formula) {
        $selection = $word.Selection
        $selection.EndKey(6) | Out-Null
        if ($doc.Content.End -gt 1) {
            $selection.TypeParagraph()
        }
        $start = $selection.Start
        $wrapped = '$' + $tex + '$'
        $selection.TypeText($wrapped)
        $end = $selection.Start
        $range = $doc.Range($start, $end)
        $range.Select()
        $beforeCount = $doc.InlineShapes.Count
        $word.Run("MTCommand_TeXToggle")
        $createdShape = $null
        for ($attempt = 0; $attempt -lt 40; $attempt++) {
            if ($doc.InlineShapes.Count -eq ($beforeCount + 1)) {
                $candidate = $doc.InlineShapes.Item($beforeCount + 1)
                $progId = ""
                try {
                    $progId = $candidate.OLEFormat.ProgID
                } catch {
                    $progId = ""
                }
                if ($progId -eq "Equation.DSMT4") {
                    $createdShape = $candidate
                    break
                }
            }
            Start-Sleep -Milliseconds 100
        }
        if ($null -eq $createdShape) {
            throw "MathType TeXToggle did not create exactly one Equation.DSMT4 object for formula: $tex"
        }
        $mathTypeOleCount++
        $equations += [PSCustomObject]@{
            docObjectIndex = $doc.InlineShapes.Count
            inlineShapeIndex = $doc.InlineShapes.Count
            formulaIndex = $mathTypeOleCount - 1
            output = $tex
            status = "word-mathtype-tex-toggle"
            progId = "Equation.DSMT4"
            widthPt = [Math]::Round([double]$createdShape.Width, 3)
            heightPt = [Math]::Round([double]$createdShape.Height, 3)
        }
    }
    if ($doc.InlineShapes.Count -ne $Formula.Count) {
        throw "Expected $($Formula.Count) InlineShapes, found $($doc.InlineShapes.Count)."
    }
    if ($mathTypeOleCount -ne $Formula.Count) {
        throw "Expected $($Formula.Count) Equation.DSMT4 objects, found $mathTypeOleCount."
    }

    $doc.SaveAs([ref]$resolvedDocx, [ref]16)
    [PSCustomObject]@{
        sourcePath = $resolvedDocx
        generator = "generate_word_mathtype_tex_reference.ps1"
        equations = $equations
    } | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $resolvedReport -Encoding UTF8

    [PSCustomObject]@{
        Docx = $resolvedDocx
        SourceReport = $resolvedReport
        FormulaCount = $Formula.Count
        InlineShapeCount = $doc.InlineShapes.Count
        MathTypeOleCount = $mathTypeOleCount
    }
} finally {
    if ($doc) {
        $doc.Close($false) | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) | Out-Null
    }
    if ($word) {
        $word.Quit() | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
    }
}
