param(
    [Parameter(Mandatory = $true)]
    [int]$WordProcessId,
    [Parameter(Mandatory = $true)]
    [string]$ResultPath,
    [ValidateSet("WholeDocument", "Selection")]
    [string]$FormatScope = "WholeDocument",
    [ValidateSet("NewEquation", "Clipboard")]
    [string]$PreferenceSource = "NewEquation",
    [int]$TimeoutSeconds = 120
)

$ErrorActionPreference = "Stop"
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName System.Windows.Forms
if (-not ("MathTypePreferenceDialogNative" -as [type])) {
Add-Type @'
using System;
using System.Runtime.InteropServices;
public static class MathTypePreferenceDialogNative {
    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
    public delegate bool EnumChildProc(IntPtr hWnd, IntPtr lParam);
    [StructLayout(LayoutKind.Sequential)]
    private struct RECT { public int Left, Top, Right, Bottom; }
    [DllImport("user32.dll")]
    public static extern bool SetForegroundWindow(IntPtr hWnd);
    [DllImport("user32.dll")]
    private static extern bool EnumWindows(EnumWindowsProc callback, IntPtr lParam);
    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
    [DllImport("user32.dll")]
    private static extern bool IsWindowVisible(IntPtr hWnd);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetClassName(IntPtr hWnd, System.Text.StringBuilder className, int maxCount);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetWindowText(IntPtr hWnd, System.Text.StringBuilder text, int maxCount);
    [DllImport("user32.dll")]
    private static extern bool EnumChildWindows(IntPtr parent, EnumChildProc callback, IntPtr lParam);
    [DllImport("user32.dll")]
    private static extern bool GetWindowRect(IntPtr hWnd, out RECT rect);
    [DllImport("user32.dll")]
    private static extern bool PostMessage(IntPtr hWnd, uint message, IntPtr wParam, IntPtr lParam);
    [DllImport("user32.dll")]
    private static extern IntPtr SendMessage(IntPtr hWnd, uint message, IntPtr wParam, IntPtr lParam);
    [DllImport("user32.dll")]
    private static extern int GetDlgCtrlID(IntPtr hWnd);

    private static string WindowClass(IntPtr hWnd) {
        var value = new System.Text.StringBuilder(128);
        GetClassName(hWnd, value, value.Capacity);
        return value.ToString();
    }

    public static string GetWindowClass(IntPtr hWnd) { return WindowClass(hWnd); }

    public static IntPtr FindVisibleDialog(int processId) {
        IntPtr found = IntPtr.Zero;
        EnumWindows(delegate(IntPtr hWnd, IntPtr lParam) {
            uint ownerProcessId;
            GetWindowThreadProcessId(hWnd, out ownerProcessId);
            if (ownerProcessId != (uint)processId || !IsWindowVisible(hWnd)) return true;
            string value = WindowClass(hWnd);
            if (value == "ThunderDFrame" || value == "#32770") {
                found = hWnd;
                return false;
            }
            return true;
        }, IntPtr.Zero);
        return found;
    }

    private static IntPtr MousePoint(int x, int y) {
        return (IntPtr)((y << 16) | (x & 0xffff));
    }

    private static void ClickAt(IntPtr hWnd, int x, int y) {
        IntPtr point = MousePoint(x, y);
        PostMessage(hWnd, 0x0201, (IntPtr)1, point);
        PostMessage(hWnd, 0x0202, IntPtr.Zero, point);
        System.Threading.Thread.Sleep(80);
    }

    private static void Key(IntPtr hWnd, int virtualKey) {
        PostMessage(hWnd, 0x0100, (IntPtr)virtualKey, IntPtr.Zero);
        PostMessage(hWnd, 0x0101, (IntPtr)virtualKey, IntPtr.Zero);
        System.Threading.Thread.Sleep(80);
    }

    public static bool ConfigureAndAccept(IntPtr dialog, bool selection, bool clipboardPreferences) {
        if (selection) {
            if (clipboardPreferences) {
                IntPtr selectionPreferences = IntPtr.Zero;
                EnumChildWindows(dialog, delegate(IntPtr child, IntPtr lParam) {
                    if (!WindowClass(child).StartsWith("F3 Server")) return true;
                    RECT rect;
                    if (!GetWindowRect(child, out rect)) return true;
                    int width = rect.Right - rect.Left;
                    int height = rect.Bottom - rect.Top;
                    if (height >= 120 && width < 420) selectionPreferences = child;
                    return true;
                }, IntPtr.Zero);
                if (selectionPreferences == IntPtr.Zero) return false;
                ClickAt(selectionPreferences, 10, 70);
                uint clipboardProcessId;
                GetWindowThreadProcessId(dialog, out clipboardProcessId);
                System.Threading.Thread.Sleep(1000);
                IntPtr clipboardRebuilt = FindVisibleDialog((int)clipboardProcessId);
                Key(clipboardRebuilt != IntPtr.Zero ? clipboardRebuilt : dialog, 0x0D);
                return true;
            }
            // In selection mode the VBA form rebuilds its windowless controls
            // after a preference click, invalidating cached F3 child handles.
            // Default focus is OK; three previous-dialog moves reach the sole
            // "MathType new equation preferences" option.
            for (int i = 0; i < 3; i++) {
                SendMessage(dialog, 0x0028, (IntPtr)1, IntPtr.Zero); // WM_NEXTDLGCTL previous
            }
            Key(dialog, 0x20); // SPACE selects the preference
            uint processId;
            GetWindowThreadProcessId(dialog, out processId);
            System.Threading.Thread.Sleep(1000); // the VBA form rebuilds after preference selection
            IntPtr rebuilt = FindVisibleDialog((int)processId);
            Key(rebuilt != IntPtr.Zero ? rebuilt : dialog, 0x0D); // invoke the rebuilt form's default OK
            return true;
        }

        IntPtr main = IntPtr.Zero;
        IntPtr preferences = IntPtr.Zero;
        IntPtr range = IntPtr.Zero;
        int mainArea = -1;
        EnumChildWindows(dialog, delegate(IntPtr child, IntPtr lParam) {
            if (!WindowClass(child).StartsWith("F3 Server")) return true;
            RECT rect;
            if (!GetWindowRect(child, out rect)) return true;
            int width = rect.Right - rect.Left;
            int height = rect.Bottom - rect.Top;
            int area = width * height;
            if (area > mainArea) { mainArea = area; main = child; }
            if (height >= 120 && width < 420) preferences = child;
            if (height >= 45 && height <= 100 && width < 420) range = child;
            return true;
        }, IntPtr.Zero);
        if (main == IntPtr.Zero || preferences == IntPtr.Zero || range == IntPtr.Zero) return false;

        // Select the pinned preference source explicitly.
        ClickAt(preferences, clipboardPreferences ? 10 : 18, clipboardPreferences ? 70 : 49);
        RECT rangeRect;
        GetWindowRect(range, out rangeRect);
        int rangeHeight = rangeRect.Bottom - rangeRect.Top;
        ClickAt(range, 18, 28);

        uint dialogProcessId;
        GetWindowThreadProcessId(dialog, out dialogProcessId);
        SetForegroundWindow(dialog);
        Key(dialog, 0x0D);
        System.Threading.Thread.Sleep(300);
        IntPtr next = FindVisibleDialog((int)dialogProcessId);
        if (next == dialog) {
            RECT mainRect;
            GetWindowRect(main, out mainRect);
            ClickAt(main, (mainRect.Right - mainRect.Left) - 44, 52);
            System.Threading.Thread.Sleep(300);
            next = FindVisibleDialog((int)dialogProcessId);
        }
        return next == IntPtr.Zero || WindowClass(next) != "ThunderDFrame";
    }

    public static string ReadDialogText(IntPtr dialog) {
        var result = new System.Text.StringBuilder();
        EnumChildWindows(dialog, delegate(IntPtr child, IntPtr lParam) {
            if (WindowClass(child) != "Static") return true;
            var text = new System.Text.StringBuilder(1024);
            GetWindowText(child, text, text.Capacity);
            if (text.Length > 0) result.AppendLine(text.ToString());
            return true;
        }, IntPtr.Zero);
        return result.ToString();
    }

    public static bool DismissStatistics(IntPtr dialog) {
        IntPtr button = IntPtr.Zero;
        EnumChildWindows(dialog, delegate(IntPtr child, IntPtr lParam) {
            if (WindowClass(child) == "Button" && GetDlgCtrlID(child) == 2) {
                button = child;
                return false;
            }
            return true;
        }, IntPtr.Zero);
        if (button == IntPtr.Zero) return false;
        SendMessage(button, 0x00F5, IntPtr.Zero, IntPtr.Zero); // BM_CLICK
        return true;
    }
}
'@
}

function Find-WordWindow {
    $dialogHandle = [MathTypePreferenceDialogNative]::FindVisibleDialog($WordProcessId)
    if ($dialogHandle -ne [IntPtr]::Zero) {
        return [System.Windows.Automation.AutomationElement]::FromHandle($dialogHandle)
    }
    $root = [System.Windows.Automation.AutomationElement]::RootElement
    $condition = New-Object System.Windows.Automation.PropertyCondition(
        [System.Windows.Automation.AutomationElement]::ProcessIdProperty,
        $WordProcessId)
    $windows = $root.FindAll([System.Windows.Automation.TreeScope]::Children, $condition)
    $mainWindow = $null
    foreach ($window in $windows) {
        if ($window.Current.ClassName -eq "OpusApp") {
            $mainWindow = $window
            continue
        }
        # MathType's VBA user forms are owned top-level windows (ThunderDFrame),
        # not descendants of Word's OpusApp UIA node. Prefer the visible modal
        # so its radio buttons and OK button are discoverable reliably.
        if (-not $window.Current.IsOffscreen -and
            $window.Current.IsEnabled -and
            -not [string]::IsNullOrWhiteSpace($window.Current.Name)) {
            return $window
        }
    }
    return $mainWindow
}

function Find-Control {
    param(
        [System.Windows.Automation.AutomationElement]$Root,
        [System.Windows.Automation.ControlType]$Type,
        [string[]]$Names
    )
    $typeCondition = New-Object System.Windows.Automation.PropertyCondition(
        [System.Windows.Automation.AutomationElement]::ControlTypeProperty,
        $Type)
    $controls = $Root.FindAll([System.Windows.Automation.TreeScope]::Descendants, $typeCondition)
    foreach ($control in $controls) {
        if ($control.Current.IsEnabled -and $Names -contains $control.Current.Name) {
            return $control
        }
    }
    return $null
}

function Invoke-Control {
    param([System.Windows.Automation.AutomationElement]$Control)
    $pattern = $Control.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern)
    $pattern.Invoke()
}

function Send-DialogKey {
    param([string]$Key)
    $shell = New-Object -ComObject WScript.Shell
    if (-not $shell.AppActivate($WordProcessId)) {
        throw "could not activate Word process $WordProcessId"
    }
    Start-Sleep -Milliseconds 100
    $shell.SendKeys($Key)
}

function Find-ControlByAutomationId {
    param(
        [System.Windows.Automation.AutomationElement]$Root,
        [System.Windows.Automation.ControlType]$Type,
        [string]$AutomationId
    )
    $typeCondition = New-Object System.Windows.Automation.PropertyCondition(
        [System.Windows.Automation.AutomationElement]::ControlTypeProperty,
        $Type)
    $controls = $Root.FindAll([System.Windows.Automation.TreeScope]::Descendants, $typeCondition)
    foreach ($control in $controls) {
        if ($control.Current.IsEnabled -and $control.Current.AutomationId -eq $AutomationId) {
            return $control
        }
    }
    return $null
}

function Select-Control {
    param([System.Windows.Automation.AutomationElement]$Control)
    $pattern = $Control.GetCurrentPattern([System.Windows.Automation.SelectionItemPattern]::Pattern)
    if (-not $pattern.Current.IsSelected) { $pattern.Select() }
}

$deadline = (Get-Date).AddSeconds($TimeoutSeconds)
$dialogAccepted = $false
$statisticsDismissed = $false
$dialogClosed = $false
$formattedCount = $null
$statisticsText = ""
$lastAutomationError = ""
$currentSelectionNames = @(
    [string]::Concat([char[]]@(0x5F53, 0x524D, 0x6240, 0x9009, 0x5185, 0x5BB9)),
    [string]::Concat([char[]]@(0x5F53, 0x524D, 0x9009, 0x5B9A, 0x5185, 0x5BB9)),
    [string]::Concat([char[]]@(0x5F53, 0x524D, 0x9009, 0x62E9, 0x5185, 0x5BB9)),
    "Current selection",
    "Selection"
)
$wholeDocumentNames = @(
    [string]::Concat([char[]]@(0x6574, 0x7BC7, 0x6587, 0x6863)),
    "Whole document"
)
$clipboardPreferenceNames = @(
    [string]::Concat([char[]]@(0x526A, 0x8D34, 0x677F, 0x4E0A, 0x7684, 0x516C, 0x5F0F)),
    "Equation on clipboard"
)
$newEquationPreferenceNames = @(
    "MathType " + [string]::Concat([char[]]@(0x65B0, 0x7684, 0x516C, 0x5F0F, 0x9884, 0x7F6E)),
    "New equation preferences"
)
$acceptNames = @(
    [string]::Concat([char[]]@(0x786E, 0x5B9A)),
    "OK"
)

while ((Get-Date) -lt $deadline -and -not $dialogClosed) {
    try {
        $nativeDialog = [MathTypePreferenceDialogNative]::FindVisibleDialog($WordProcessId)
        if ($nativeDialog -ne [IntPtr]::Zero) {
            $nativeClass = [MathTypePreferenceDialogNative]::GetWindowClass($nativeDialog)
            if (-not $dialogAccepted -and $nativeClass -eq "ThunderDFrame") {
                if ([MathTypePreferenceDialogNative]::ConfigureAndAccept(
                        $nativeDialog, $FormatScope -eq "Selection", $PreferenceSource -eq "Clipboard")) {
                    $dialogAccepted = $true
                }
                Start-Sleep -Milliseconds 100
                continue
            }
            if ($dialogAccepted -and -not $statisticsDismissed -and $nativeClass -eq "#32770") {
                $statisticsText = [MathTypePreferenceDialogNative]::ReadDialogText($nativeDialog)
                if ($statisticsText -match "(?<count>\d+)\s+MathType") {
                    $formattedCount = [int]$Matches.count
                    if ([MathTypePreferenceDialogNative]::DismissStatistics($nativeDialog)) {
                        $statisticsDismissed = $true
                    }
                }
                Start-Sleep -Milliseconds 100
                continue
            }
        } elseif ($dialogAccepted -and $statisticsDismissed) {
            $dialogClosed = $true
            continue
        }

        $window = Find-WordWindow
        if (-not $window) {
            Start-Sleep -Milliseconds 100
            continue
        }

        if (-not $dialogAccepted) {
            $scopeNames = if ($FormatScope -eq "Selection") {
                $currentSelectionNames
            } else {
                $wholeDocumentNames
            }
            $scopeControl = Find-Control $window ([System.Windows.Automation.ControlType]::RadioButton) $scopeNames
            $preferenceNames = if ($PreferenceSource -eq "Clipboard") {
                $clipboardPreferenceNames
            } else {
                $newEquationPreferenceNames
            }
            $newEquationPreferences = Find-Control $window ([System.Windows.Automation.ControlType]::RadioButton) $preferenceNames
            $accept = Find-Control $window ([System.Windows.Automation.ControlType]::Button) $acceptNames
            if ($scopeControl -and $newEquationPreferences -and $accept) {
            Select-Control $scopeControl
            Select-Control $newEquationPreferences
            $dialogAccepted = $true
            Send-DialogKey "{ENTER}"
            }
            Start-Sleep -Milliseconds 100
            continue
        }

        if (-not $statisticsDismissed) {
            $textCondition = New-Object System.Windows.Automation.PropertyCondition(
                [System.Windows.Automation.AutomationElement]::ControlTypeProperty,
                [System.Windows.Automation.ControlType]::Text)
            $texts = $window.FindAll(
                [System.Windows.Automation.TreeScope]::Descendants,
                $textCondition)
            foreach ($text in $texts) {
                if ($text.Current.Name -match "MathType" -and
                    $text.Current.Name -match "(?<count>\d+)") {
                    $statisticsText = $text.Current.Name
                    $formattedCount = [int]$Matches.count
                    break
                }
            }
            if ($null -ne $formattedCount) {
                $dismiss = Find-ControlByAutomationId $window ([System.Windows.Automation.ControlType]::Button) "2"
                if ($dismiss) {
                    $statisticsDismissed = $true
                    Send-DialogKey "{ENTER}"
                }
            }
            Start-Sleep -Milliseconds 100
            continue
        }

        $dialogClosed = $true
        Send-DialogKey "{ESC}"
    } catch {
        # Word rebuilds parts of its UI tree while the MathType macro runs.
        # Retry transient UI Automation failures until the overall deadline.
        $lastAutomationError = $_.Exception.ToString()
    }
    if (-not $dialogClosed) {
        Start-Sleep -Milliseconds 100
    }
}

$result = [ordered]@{
    schemaVersion = 1
    dialogAccepted = $dialogAccepted
    statisticsDismissed = $statisticsDismissed
    dialogClosed = $dialogClosed
    formattedObjectCount = $formattedCount
    statisticsText = $statisticsText
    preferenceSource = $PreferenceSource
    lastAutomationError = $lastAutomationError
}
$result | ConvertTo-Json -Depth 3 | Set-Content -LiteralPath $ResultPath -Encoding UTF8

if (-not $dialogClosed) {
    throw "MathType Format Equations dialog did not complete within $TimeoutSeconds seconds; last error: $lastAutomationError"
}
