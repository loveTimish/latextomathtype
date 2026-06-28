param(
    [Parameter(Mandatory = $true)]
    [string]$DocxPath,
    [string]$OutDir = "",
    [int]$SpotcheckFormulaIndex = -1,
    [int]$OpenDelaySeconds = 8,
    [switch]$NoMathTypeUi,
    [switch]$KeepOpen
)

$ErrorActionPreference = "Stop"

$resolvedDocx = Resolve-Path -LiteralPath $DocxPath
if ([string]::IsNullOrWhiteSpace($OutDir)) {
    $leaf = [IO.Path]::GetFileNameWithoutExtension($resolvedDocx.Path)
    $OutDir = Join-Path "analysis\word-mathtype-reference" $leaf
}
$resolvedOut = [IO.Path]::GetFullPath((Join-Path (Get-Location) $OutDir))
New-Item -ItemType Directory -Force -Path $resolvedOut | Out-Null

Add-Type @"
using System;
using System.Runtime.InteropServices;
public static class WordMtRefWin32 {
    public const int HWND_TOPMOST = -1;
    public const int HWND_NOTOPMOST = -2;
    public const int SWP_NOSIZE = 0x0001;
    public const int SWP_NOMOVE = 0x0002;
    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
    [DllImport("user32.dll")]
    public static extern bool SetForegroundWindow(IntPtr hWnd);
    [DllImport("user32.dll")]
    public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);
    [DllImport("user32.dll")]
    public static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);
    [DllImport("user32.dll")]
    public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
    [DllImport("user32.dll")]
    public static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);
    [DllImport("user32.dll")]
    public static extern bool IsWindowVisible(IntPtr hWnd);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)]
    public static extern int GetWindowText(IntPtr hWnd, System.Text.StringBuilder lpString, int nMaxCount);
    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
    public struct RECT {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }
}
"@

function Get-VisibleWindows {
    $windows = New-Object System.Collections.Generic.List[object]
    $callback = [WordMtRefWin32+EnumWindowsProc]{
        param([IntPtr]$hWnd, [IntPtr]$lParam)
        if (-not [WordMtRefWin32]::IsWindowVisible($hWnd)) {
            return $true
        }
        $titleBuilder = New-Object System.Text.StringBuilder 512
        [void][WordMtRefWin32]::GetWindowText($hWnd, $titleBuilder, $titleBuilder.Capacity)
        $title = $titleBuilder.ToString()
        if ([string]::IsNullOrWhiteSpace($title)) {
            return $true
        }
        [uint32]$processId = 0
        [void][WordMtRefWin32]::GetWindowThreadProcessId($hWnd, [ref]$processId)
        $windows.Add([PSCustomObject]@{
            Hwnd = $hWnd
            Pid = [int]$processId
            Title = $title
        })
        return $true
    }
    [void][WordMtRefWin32]::EnumWindows($callback, [IntPtr]::Zero)
    return $windows
}

function Find-MathTypeWindow([string]$DocxLeaf) {
    $windows = @(Get-VisibleWindows)
    $matchingDoc = $windows | Where-Object {
        $_.Title -like "MathType*" -and $_.Title -like "*$DocxLeaf*"
    } | Select-Object -First 1
    if ($matchingDoc) {
        return $matchingDoc
    }
    return $windows | Where-Object { $_.Title -like "MathType*" } | Select-Object -First 1
}

function Set-TopMost([IntPtr]$Hwnd, [bool]$Enabled) {
    if ($Hwnd -eq [IntPtr]::Zero) {
        return
    }
    $after = if ($Enabled) { [IntPtr][WordMtRefWin32]::HWND_TOPMOST } else { [IntPtr][WordMtRefWin32]::HWND_NOTOPMOST }
    [WordMtRefWin32]::SetWindowPos(
        $Hwnd,
        $after,
        0,
        0,
        0,
        0,
        [uint32]([WordMtRefWin32]::SWP_NOMOVE -bor [WordMtRefWin32]::SWP_NOSIZE)
    ) | Out-Null
}

function Capture-Window([string]$Path, [IntPtr]$WindowHandle) {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    $rect = New-Object WordMtRefWin32+RECT
    if (-not [WordMtRefWin32]::GetWindowRect($WindowHandle, [ref]$rect)) {
        throw "Could not read window rectangle for screenshot."
    }
    $width = [Math]::Max(1, $rect.Right - $rect.Left)
    $height = [Math]::Max(1, $rect.Bottom - $rect.Top)
    $bitmap = New-Object System.Drawing.Bitmap $width, $height
    $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
    try {
        $graphics.CopyFromScreen($rect.Left, $rect.Top, 0, 0, (New-Object System.Drawing.Size $width, $height))
        $bitmap.Save($Path, [System.Drawing.Imaging.ImageFormat]::Png)
    } finally {
        $graphics.Dispose()
        $bitmap.Dispose()
    }
}

$jsonPath = Join-Path $resolvedOut "word-mathtype-reference.json"
$pdfPath = Join-Path $resolvedOut "word-reference.pdf"
$spotcheckPath = if ($SpotcheckFormulaIndex -ge 0) {
    Join-Path $resolvedOut ("mathtype-formula-{0}.png" -f $SpotcheckFormulaIndex)
} else {
    ""
}

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.ScreenUpdating = $false
$word.DisplayAlerts = 0
$doc = $null
try {
    $doc = $word.Documents.Open($resolvedDocx.Path, $false, $false)
    $items = @()
    $oleOrdinal = -1
    $target = $null
    for ($i = 1; $i -le $doc.InlineShapes.Count; $i++) {
        $shape = $doc.InlineShapes.Item($i)
        $progId = $null
        try {
            $progId = $shape.OLEFormat.ProgID
        } catch {
            $progId = $null
        }
        if (-not $progId) {
            continue
        }
        if ($progId -eq "Equation.DSMT4") {
            $oleOrdinal++
            if ($oleOrdinal -eq $SpotcheckFormulaIndex) {
                $target = $shape
            }
        }
        $range = $shape.Range
        $items += [PSCustomObject]@{
            inlineShapeIndex = $i
            formulaIndex = if ($progId -eq "Equation.DSMT4") { $oleOrdinal } else { $null }
            progId = $progId
            widthPt = [Math]::Round([double]$shape.Width, 3)
            heightPt = [Math]::Round([double]$shape.Height, 3)
            linePositionHalfPt = $range.Font.Position
            paragraphIndex = $range.Paragraphs.Item(1).Range.Start
            rangeStart = $range.Start
            rangeEnd = $range.End
        }
    }

    $doc.ExportAsFixedFormat($pdfPath, 17)

    if ($SpotcheckFormulaIndex -ge 0 -and -not $NoMathTypeUi) {
        if ($null -eq $target) {
            throw "Could not find Equation.DSMT4 formula index $SpotcheckFormulaIndex."
        }
        $target.Range.Select()
        $word.Activate()
        try {
            $target.OLEFormat.Edit()
        } catch {
            try {
                $target.OLEFormat.DoVerb(-2)
            } catch {
                $target.OLEFormat.DoVerb(0)
            }
        }
        Start-Sleep -Seconds $OpenDelaySeconds
        $mathTypeWindow = Find-MathTypeWindow ([IO.Path]::GetFileName($resolvedDocx.Path))
        if ($null -eq $mathTypeWindow) {
            throw "Could not find a visible MathType window."
        }
        [WordMtRefWin32]::ShowWindowAsync($mathTypeWindow.Hwnd, 9) | Out-Null
        [WordMtRefWin32]::SetForegroundWindow($mathTypeWindow.Hwnd) | Out-Null
        Set-TopMost $mathTypeWindow.Hwnd $true
        Start-Sleep -Milliseconds 500
        Capture-Window $spotcheckPath $mathTypeWindow.Hwnd
        Set-TopMost $mathTypeWindow.Hwnd $false
    } elseif ($SpotcheckFormulaIndex -ge 0 -and $NoMathTypeUi) {
        $spotcheckPath = ""
    }

    [PSCustomObject]@{
        docx = $resolvedDocx.Path
        pdf = $pdfPath
        screenshot = $spotcheckPath
        inlineShapeCount = $doc.InlineShapes.Count
        mathTypeOleCount = ($items | Where-Object { $_.progId -eq "Equation.DSMT4" }).Count
        items = $items
    } | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $jsonPath -Encoding UTF8

    [PSCustomObject]@{
        Docx = $resolvedDocx.Path
        Json = $jsonPath
        Pdf = $pdfPath
        Screenshot = $spotcheckPath
        MathTypeOleCount = ($items | Where-Object { $_.progId -eq "Equation.DSMT4" }).Count
        InlineShapeCount = $doc.InlineShapes.Count
    }
} finally {
    if (-not $KeepOpen -and $doc -ne $null) {
        $doc.Close($false)
    }
    if (-not $KeepOpen) {
        $word.Quit()
    }
}
