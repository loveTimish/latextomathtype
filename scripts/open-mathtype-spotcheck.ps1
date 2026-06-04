param(
    [string]$DocxPath = "target\reference-roundtrip\fraction-split-reference-regenerated-shell.docx",
    [int]$FormulaIndex = 9,
    [string]$OutPng = "",
    [int]$OpenDelaySeconds = 12,
    [switch]$KeepOpen
)

$ErrorActionPreference = "Stop"

if ([string]::IsNullOrWhiteSpace($OutPng)) {
    $OutPng = "target\reference-roundtrip\mathtype-spotcheck-formula-$FormulaIndex.png"
}

$resolvedDocx = Resolve-Path -LiteralPath $DocxPath
$resolvedOut = [IO.Path]::GetFullPath((Join-Path (Get-Location) $OutPng))
New-Item -ItemType Directory -Force -Path ([IO.Path]::GetDirectoryName($resolvedOut)) | Out-Null

Add-Type @"
using System;
using System.Runtime.InteropServices;
public static class Win32Window {
    public const int HWND_TOPMOST = -1;
    public const int HWND_NOTOPMOST = -2;
    public const int SWP_NOSIZE = 0x0001;
    public const int SWP_NOMOVE = 0x0002;
    public const int SWP_NOACTIVATE = 0x0010;
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

function Show-ProcessWindow([System.Diagnostics.Process]$Process) {
    if ($null -eq $Process -or $Process.MainWindowHandle -eq [IntPtr]::Zero) {
        return
    }
    Show-Hwnd $Process.MainWindowHandle
}

function Show-Hwnd([IntPtr]$Hwnd) {
    if ($Hwnd -eq [IntPtr]::Zero) {
        return
    }
    [Win32Window]::ShowWindowAsync($Hwnd, 9) | Out-Null
    [Win32Window]::SetForegroundWindow($Hwnd) | Out-Null
}

function Set-TopMost([IntPtr]$Hwnd, [bool]$Enabled) {
    if ($Hwnd -eq [IntPtr]::Zero) {
        return
    }
    $after = if ($Enabled) { [IntPtr][Win32Window]::HWND_TOPMOST } else { [IntPtr][Win32Window]::HWND_NOTOPMOST }
    [Win32Window]::SetWindowPos(
        $Hwnd,
        $after,
        0,
        0,
        0,
        0,
        [uint32]([Win32Window]::SWP_NOMOVE -bor [Win32Window]::SWP_NOSIZE)
    ) | Out-Null
}

function Get-VisibleWindows {
    $windows = New-Object System.Collections.Generic.List[object]
    $callback = [Win32Window+EnumWindowsProc]{
        param([IntPtr]$hWnd, [IntPtr]$lParam)
        if (-not [Win32Window]::IsWindowVisible($hWnd)) {
            return $true
        }
        $titleBuilder = New-Object System.Text.StringBuilder 512
        [void][Win32Window]::GetWindowText($hWnd, $titleBuilder, $titleBuilder.Capacity)
        $title = $titleBuilder.ToString()
        if ([string]::IsNullOrWhiteSpace($title)) {
            return $true
        }
        [uint32]$processId = 0
        [void][Win32Window]::GetWindowThreadProcessId($hWnd, [ref]$processId)
        $windows.Add([PSCustomObject]@{
            Hwnd = $hWnd
            Pid = [int]$processId
            Title = $title
        })
        return $true
    }
    [void][Win32Window]::EnumWindows($callback, [IntPtr]::Zero)
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
    $warningDialog = $windows | Where-Object {
        $_.Title -like "*MathType*" -and $_.Title -like "*警告*"
    } | Select-Object -First 1
    if ($warningDialog) {
        return $warningDialog
    }
    $anyMathType = $windows | Where-Object { $_.Title -like "MathType*" } | Select-Object -First 1
    if ($anyMathType) {
        return $anyMathType
    }
    return $null
}

function Capture-Screen([string]$Path, [IntPtr]$WindowHandle = [IntPtr]::Zero) {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    $bounds = [System.Windows.Forms.SystemInformation]::VirtualScreen
    if ($WindowHandle -ne [IntPtr]::Zero) {
        $rect = New-Object Win32Window+RECT
        if ([Win32Window]::GetWindowRect($WindowHandle, [ref]$rect)) {
            $width = [Math]::Max(1, $rect.Right - $rect.Left)
            $height = [Math]::Max(1, $rect.Bottom - $rect.Top)
            $bounds = New-Object System.Drawing.Rectangle $rect.Left, $rect.Top, $width, $height
        }
    }
    $bitmap = New-Object System.Drawing.Bitmap $bounds.Width, $bounds.Height
    $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
    try {
        $graphics.CopyFromScreen($bounds.Left, $bounds.Top, 0, 0, $bounds.Size)
        $bitmap.Save($Path, [System.Drawing.Imaging.ImageFormat]::Png)
    } finally {
        $graphics.Dispose()
        $bitmap.Dispose()
    }
}

$word = New-Object -ComObject Word.Application
$word.Visible = $true
$doc = $null
try {
    $doc = $word.Documents.Open($resolvedDocx.Path, $false, $false)
    $word.Activate()
    Start-Sleep -Milliseconds 500
    $target = $null
    $oleOrdinal = -1
    for ($i = 1; $i -le $doc.InlineShapes.Count; $i++) {
        $shape = $doc.InlineShapes.Item($i)
        try {
            $progId = $shape.OLEFormat.ProgID
        } catch {
            continue
        }
        if ($progId -ne "Equation.DSMT4") {
            continue
        }
        $oleOrdinal++
        if ($oleOrdinal -eq $FormulaIndex) {
            $target = $shape
            break
        }
    }

    if ($null -eq $target) {
        throw "Could not find MathType OLE formula index $FormulaIndex in $($resolvedDocx.Path)"
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

    $docxLeaf = [IO.Path]::GetFileName($resolvedDocx.Path)
    $mathTypeWindow = Find-MathTypeWindow $docxLeaf
    if ($mathTypeWindow) {
        Show-Hwnd $mathTypeWindow.Hwnd
        Set-TopMost $mathTypeWindow.Hwnd $true
        Start-Sleep -Milliseconds 500
    } else {
        throw "Could not find a visible MathType window after opening formula index $FormulaIndex"
    }
    Capture-Screen $resolvedOut $mathTypeWindow.Hwnd
    Set-TopMost $mathTypeWindow.Hwnd $false

    [PSCustomObject]@{
        Docx = $resolvedDocx.Path
        FormulaIndex = $FormulaIndex
        ProgID = "Equation.DSMT4"
        Screenshot = $resolvedOut
        MathTypeWindow = if ($mathTypeWindow) { $mathTypeWindow.Title } else { "" }
    }
} finally {
    if (-not $KeepOpen -and $doc -ne $null) {
        $doc.Close($false)
    }
    if (-not $KeepOpen) {
        $word.Quit()
    }
}
