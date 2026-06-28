# Export Word DOCX to PDF for visual AI auditing
param(
    [Parameter(Mandatory=$true)]
    [string]$DocxPath,
    
    [string]$OutDir = $null,
    
    [int]$Dpi = 200
)

$ErrorActionPreference = "Stop"

# Resolve paths
$docxFile = (Resolve-Path -LiteralPath $DocxPath).Path
if (-not $OutDir) {
    $OutDir = Join-Path (Split-Path $docxFile -Parent) "visual-audit"
}

# Create output directory
New-Item -ItemType Directory -Path $OutDir -Force | Out-Null
Write-Host "📁 Output directory: $OutDir"

# Step 1: Word -> PDF using Word COM
Write-Host "`n📄 Exporting DOCX to PDF..."
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0

try {
    $missing = [System.Type]::Missing
    $doc = $word.Documents.Open($docxFile)
    $pdfPath = Join-Path $OutDir ((Get-Item $docxFile).BaseName + ".pdf")
    
    # Word SaveAs2 method with proper COM arguments
    $wdFormatPDF = 17
    $doc.SaveAs2($pdfPath, $wdFormatPDF)
    $doc.Close()
    Write-Host "✅ PDF saved: $pdfPath"
} catch {
    Write-Host "❌ PDF export failed: $_"
    try { $word.Quit() } catch {}
    exit 1
} finally {
    try { $word.Quit() } catch {}
}

Write-Host "`n🎉 Visual audit materials are ready!"
Write-Host "`n📂 Output folder: $OutDir"
Write-Host "   - PDF: $pdfPath"
Write-Host ""
Write-Host "📋 Visual AI Audit Instructions:"
Write-Host "   1. Open the PDF in Adobe Acrobat / Chrome / any viewer"
Write-Host "   2. Zoom to 100% and take screenshots of formula pages"
Write-Host "   3. Upload PNGs to GPT-4V / Claude 3 Opus / Gemini Ultra"
Write-Host "   4. Use this prompt:"
Write-Host ""
Write-Host '      "As a math typesetting expert, review these LaTeX-to-Word converted formulas:'
Write-Host '      - Check for visual correctness, symbol rendering quality'
Write-Host '      - Verify fraction bars, radical signs, and mathematical spacing'
Write-Host '      - Identify any rendering artifacts or broken symbols'
Write-Host '      - Rate overall visual quality: Excellent / Good / Fair / Poor"'
Write-Host ""

# Open the output folder
explorer.exe $OutDir
