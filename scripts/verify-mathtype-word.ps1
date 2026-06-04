param(
    [string]$DocxPath = "target\rebuild\reorganized-paper-styled.docx",
    [int]$MinimumOleCount = 1
)

$ErrorActionPreference = "Stop"

$resolvedDocx = (Resolve-Path -LiteralPath $DocxPath).Path
$word = New-Object -ComObject Word.Application
$word.Visible = $false

try {
    $doc = $word.Documents.Open($resolvedDocx, $false, $true)
    try {
        $oleItems = @()
        for ($i = 1; $i -le $doc.InlineShapes.Count; $i++) {
            $shape = $doc.InlineShapes.Item($i)
            $progId = $null
            try {
                $progId = $shape.OLEFormat.ProgID
            } catch {
                $progId = $null
            }
            if ($progId) {
                $oleItems += [PSCustomObject]@{
                    Index = $i
                    Type = $shape.Type
                    ProgID = $progId
                    Width = [double]$shape.Width
                    Height = [double]$shape.Height
                }
            }
        }

        if ($oleItems.Count -lt $MinimumOleCount) {
            throw "Expected at least $MinimumOleCount OLE objects, found $($oleItems.Count)"
        }

        $nonMathType = $oleItems | Where-Object { $_.ProgID -ne "Equation.DSMT4" }
        if ($nonMathType.Count -gt 0) {
            throw "Found non-MathType OLE ProgID: $($nonMathType[0].ProgID)"
        }

        $oleItems | Select-Object -First 30 | Format-Table -AutoSize
        Write-Output ("OLE_COUNT=" + $oleItems.Count)
        Write-Output ("INLINE_COUNT=" + $doc.InlineShapes.Count)
        Write-Output "mathtype word verification ok"
    } finally {
        $doc.Close($false)
    }
} finally {
    $word.Quit()
}
