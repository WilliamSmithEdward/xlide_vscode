param(
    [Parameter(Mandatory = $true)]
    [string]$CasePath,

    [Parameter(Mandatory = $true)]
    [string]$PidPath,

    [Parameter(Mandatory = $false)]
    [string]$StagePath = "",

    [Parameter(Mandatory = $false)]
    [string]$DialogPath = "",

    [Parameter(Mandatory = $false)]
    [int]$DialogWatchSeconds = 60
)

$ErrorActionPreference = "Stop"

Add-Type @"
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;

public static class XlideUser32 {
    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
}

public static class XlideVbeDialogWatcher {
    private const int WM_CLOSE = 0x0010;
    private const int WM_COMMAND = 0x0111;
    private const int IDOK = 1;
    private const int VBE_RESET_COMMAND_ID = 228;
    private const byte VK_MENU = 0x12;
    private const byte VK_D = 0x44;
    private const byte VK_L = 0x4C;
    private const byte VK_R = 0x52;
    private const uint KEYEVENTF_KEYUP = 0x0002;
    private static Thread watcherThread;
    private static volatile bool stopRequested;

    private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern bool EnumChildWindows(IntPtr hWndParent, EnumWindowsProc lpEnumFunc, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetWindowTextLength(IntPtr hWnd);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

    [DllImport("user32.dll")]
    private static extern bool IsWindowVisible(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern bool IsWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern bool PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern bool SetForegroundWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);

    public static void Start(uint excelProcessId, string resultPath, string stagePath, int maxSeconds) {
        Stop();
        if (excelProcessId == 0 || String.IsNullOrEmpty(resultPath)) {
            return;
        }

        stopRequested = false;
        watcherThread = new Thread(delegate() {
            Watch(excelProcessId, resultPath, stagePath, maxSeconds);
        });
        watcherThread.IsBackground = true;
        watcherThread.Start();
    }

    public static void Stop() {
        stopRequested = true;
        if (watcherThread != null && watcherThread.IsAlive) {
            watcherThread.Join(1000);
        }
        watcherThread = null;
    }

    public static bool InvokeVbeCompile(uint excelProcessId) {
        IntPtr vbeHwnd = FindVbeMainWindow(excelProcessId);
        if (vbeHwnd == IntPtr.Zero) {
            return false;
        }

        SetForegroundWindow(vbeHwnd);
        Thread.Sleep(300);
        AltTap(VK_D);
        Tap(VK_L);
        return true;
    }

    private static void Watch(uint excelProcessId, string resultPath, string stagePath, int maxSeconds) {
        DateTime deadline = DateTime.UtcNow.AddSeconds(Math.Max(1, maxSeconds));
        while (!stopRequested && DateTime.UtcNow < deadline) {
            DialogInfo dialog = FindVbeDialog(excelProcessId);
            if (dialog != null) {
                WriteDialogResult(resultPath, dialog);
                if (!String.IsNullOrEmpty(stagePath)) {
                    try {
                        string stage = dialog.Kind == "vbe_compile_dialog" ? "compile_dialog" : "vbe_dialog";
                        File.WriteAllText(stagePath, stage, Encoding.ASCII);
                    }
                    catch {
                    }
                }
                DismissDialogAndReset(dialog.Handle, excelProcessId);
                return;
            }
            Thread.Sleep(100);
        }
    }

    private static DialogInfo FindVbeDialog(uint excelProcessId) {
        DialogInfo found = null;
        EnumWindows(delegate(IntPtr hWnd, IntPtr lParam) {
            if (found != null) {
                return false;
            }

            uint windowProcessId;
            GetWindowThreadProcessId(hWnd, out windowProcessId);
            if (windowProcessId != excelProcessId || !IsWindowVisible(hWnd)) {
                return true;
            }

            string title = GetText(hWnd);
            string className = GetClass(hWnd);
            if (title != "Microsoft Visual Basic for Applications" || className != "#32770") {
                return true;
            }

            List<string> childTexts = GetChildTexts(hWnd);
            found = new DialogInfo(hWnd, title, className, childTexts);
            return false;
        }, IntPtr.Zero);
        return found;
    }

    private static IntPtr FindVbeMainWindow(uint excelProcessId) {
        IntPtr found = IntPtr.Zero;
        EnumWindows(delegate(IntPtr hWnd, IntPtr lParam) {
            if (found != IntPtr.Zero) {
                return false;
            }

            uint windowProcessId;
            GetWindowThreadProcessId(hWnd, out windowProcessId);
            if (windowProcessId != excelProcessId || !IsWindowVisible(hWnd)) {
                return true;
            }

            if (GetClass(hWnd) == "wndclass_desked_gsk") {
                found = hWnd;
                return false;
            }
            return true;
        }, IntPtr.Zero);
        return found;
    }

    private static List<string> GetChildTexts(IntPtr parent) {
        List<string> texts = new List<string>();
        EnumChildWindows(parent, delegate(IntPtr hWnd, IntPtr lParam) {
            string text = GetText(hWnd);
            if (!String.IsNullOrWhiteSpace(text)) {
                texts.Add(text.Trim());
            }
            return true;
        }, IntPtr.Zero);
        return texts;
    }

    private static string GetText(IntPtr hWnd) {
        int length = GetWindowTextLength(hWnd);
        if (length <= 0) {
            return "";
        }
        StringBuilder builder = new StringBuilder(length + 1);
        GetWindowText(hWnd, builder, builder.Capacity);
        return builder.ToString();
    }

    private static string GetClass(IntPtr hWnd) {
        StringBuilder builder = new StringBuilder(256);
        GetClassName(hWnd, builder, builder.Capacity);
        return builder.ToString();
    }

    private static void DismissDialogAndReset(IntPtr hWnd, uint excelProcessId) {
        PostMessage(hWnd, WM_COMMAND, new IntPtr(IDOK), IntPtr.Zero);
        Thread.Sleep(500);
        if (IsWindow(hWnd)) {
            PostMessage(hWnd, WM_CLOSE, IntPtr.Zero, IntPtr.Zero);
        }
        ResetVbe(excelProcessId);
    }

    private static void ResetVbe(uint excelProcessId) {
        IntPtr vbeHwnd = FindVbeMainWindow(excelProcessId);
        for (int attempt = 0; attempt < 4; attempt++) {
            EnumWindows(delegate(IntPtr hWnd, IntPtr lParam) {
                uint windowProcessId;
                GetWindowThreadProcessId(hWnd, out windowProcessId);
                if (windowProcessId != excelProcessId || !IsWindowVisible(hWnd)) {
                    return true;
                }

                string className = GetClass(hWnd);
                if (className == "wndclass_desked_gsk") {
                    PostMessage(hWnd, WM_COMMAND, new IntPtr(VBE_RESET_COMMAND_ID), IntPtr.Zero);
                }
                return true;
            }, IntPtr.Zero);
            if (vbeHwnd != IntPtr.Zero) {
                SetForegroundWindow(vbeHwnd);
                Thread.Sleep(100);
                AltTap(VK_D);
                Tap(VK_R);
            }
            Thread.Sleep(250);
        }
    }

    private static void AltTap(byte vk) {
        keybd_event(VK_MENU, 0, 0, UIntPtr.Zero);
        Thread.Sleep(50);
        Tap(vk);
        keybd_event(VK_MENU, 0, KEYEVENTF_KEYUP, UIntPtr.Zero);
        Thread.Sleep(150);
    }

    private static void Tap(byte vk) {
        keybd_event(vk, 0, 0, UIntPtr.Zero);
        Thread.Sleep(50);
        keybd_event(vk, 0, KEYEVENTF_KEYUP, UIntPtr.Zero);
        Thread.Sleep(100);
    }

    private static void WriteDialogResult(string resultPath, DialogInfo dialog) {
        StringBuilder json = new StringBuilder();
        json.Append("{");
        json.Append("\"kind\":").Append(JsonString(dialog.Kind)).Append(",");
        json.Append("\"title\":").Append(JsonString(dialog.Title)).Append(",");
        json.Append("\"className\":").Append(JsonString(dialog.ClassName)).Append(",");
        json.Append("\"message\":").Append(JsonString(dialog.Message)).Append(",");
        json.Append("\"texts\":[");
        for (int i = 0; i < dialog.Texts.Count; i++) {
            if (i > 0) {
                json.Append(",");
            }
            json.Append(JsonString(dialog.Texts[i]));
        }
        json.Append("]}");
        File.WriteAllText(resultPath, json.ToString(), Encoding.UTF8);
    }

    private static string JsonString(string value) {
        if (value == null) {
            return "null";
        }
        StringBuilder escaped = new StringBuilder();
        escaped.Append('"');
        for (int i = 0; i < value.Length; i++) {
            char ch = value[i];
            switch (ch) {
                case '\\': escaped.Append("\\\\"); break;
                case '"': escaped.Append("\\\""); break;
                case '\b': escaped.Append("\\b"); break;
                case '\f': escaped.Append("\\f"); break;
                case '\n': escaped.Append("\\n"); break;
                case '\r': escaped.Append("\\r"); break;
                case '\t': escaped.Append("\\t"); break;
                default:
                    if (ch < 32) {
                        escaped.Append("\\u").Append(((int)ch).ToString("x4"));
                    }
                    else {
                        escaped.Append(ch);
                    }
                    break;
            }
        }
        escaped.Append('"');
        return escaped.ToString();
    }

    private sealed class DialogInfo {
        public readonly IntPtr Handle;
        public readonly string Title;
        public readonly string ClassName;
        public readonly List<string> Texts;
        public readonly string Kind;
        public readonly string Message;

        public DialogInfo(IntPtr handle, string title, string className, List<string> texts) {
            Handle = handle;
            Title = title;
            ClassName = className;
            Texts = texts;
            Kind = BuildKind(texts);
            Message = BuildMessage(texts);
        }

        private static string BuildKind(List<string> texts) {
            for (int i = 0; i < texts.Count; i++) {
                if (texts[i].Contains("Compile error:")) {
                    return "vbe_compile_dialog";
                }
            }
            return "vbe_dialog";
        }

        private static string BuildMessage(List<string> texts) {
            List<string> messageParts = new List<string>();
            for (int i = 0; i < texts.Count; i++) {
                string text = texts[i];
                if (text == "OK" || text == "Help") {
                    continue;
                }
                bool alreadyAdded = false;
                for (int j = 0; j < messageParts.Count; j++) {
                    if (messageParts[j] == text) {
                        alreadyAdded = true;
                        break;
                    }
                }
                if (!alreadyAdded) {
                    messageParts.Add(text);
                }
            }
            return String.Join(" ", messageParts.ToArray());
        }
    }
}
"@

function Write-ResultJson($Value) {
    $Value | ConvertTo-Json -Depth 12 -Compress
}

function Release-ComObject($Value) {
    if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) {
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Value)
    }
}

function Set-Stage($Name) {
    $script:stage = $Name
    if ($StagePath) {
        try { Set-Content -LiteralPath $StagePath -Value $Name -Encoding ascii } catch { }
    }
}

function Read-DialogResult {
    if (-not $DialogPath -or -not (Test-Path -LiteralPath $DialogPath)) {
        return $null
    }
    try {
        return Get-Content -LiteralPath $DialogPath -Raw | ConvertFrom-Json
    }
    catch {
        return $null
    }
}

function Set-DialogOutcome($Dialog, [bool]$CompileOnly) {
    $kind = [string]$Dialog.kind
    if ($kind -eq "vbe_compile_dialog") {
        $result.outcome = "rejected"
        $result.stage = "compile_dialog"
    }
    elseif ($CompileOnly) {
        $result.outcome = "accepted"
        $result.stage = "vbe_dialog_after_compile"
    }
    else {
        $result.outcome = "rejected"
        $result.stage = "vbe_dialog"
    }
    $result.message = [string]$Dialog.message
    $result.hresult = $null
    if (-not $result.message) {
        $result.message = "VBE showed a modal dialog."
    }
}

function Wait-ForDialogResult {
    $deadline = [DateTime]::UtcNow.AddSeconds([Math]::Max(1, $DialogWatchSeconds))
    while ([DateTime]::UtcNow -lt $deadline) {
        $dialog = Read-DialogResult
        if ($null -ne $dialog) {
            return $dialog
        }
        Start-Sleep -Milliseconds 100
    }
    return $null
}

$case = Get-Content -LiteralPath $CasePath -Raw | ConvertFrom-Json
$excel = $null
$workbook = $null
$vbProject = $null
$component = $null
$codeModule = $null
$workbookPath = $null
$mode = "compile"
Set-Stage "setup"

$result = [ordered]@{
    caseId = $case.id
    outcome = "setup_error"
    stage = $stage
    message = ""
    hresult = $null
}

try {
    Set-Stage "start_excel"
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    # msoAutomationSecurityLow. This applies only to the disposable Excel
    # instance created by this worker.
    try { $excel.AutomationSecurity = 1 } catch { }

    $excelPid = 0
    [void][XlideUser32]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd, [ref]$excelPid)
    Set-Content -LiteralPath $PidPath -Value $excelPid -Encoding ascii

    Set-Stage "create_workbook"
    $workbook = $excel.Workbooks.Add()
    [void]$workbook.Activate()

    Set-Stage "add_module"
    $vbProject = $workbook.VBProject
    $component = $vbProject.VBComponents.Add(1)
    $component.Name = "XlideOracleModule"
    $codeModule = $component.CodeModule
    $codeModule.AddFromString([string]$case.source)

    if ($case.PSObject.Properties.Name -contains "mode" -and $case.mode) {
        $mode = [string]$case.mode
    }

    if ($mode -eq "compile") {
        Set-Stage "compile"
        $excel.Visible = $true
        $excel.VBE.MainWindow.Visible = $true

        [void]$component.Activate()
        $codeModule.CodePane.Show()
        $codeModule.CodePane.SetSelection(1, 1, 1, 1)
        Start-Sleep -Milliseconds 250

        [XlideVbeDialogWatcher]::Start($excelPid, $DialogPath, $StagePath, $DialogWatchSeconds)
        try {
            if (-not [XlideVbeDialogWatcher]::InvokeVbeCompile($excelPid)) {
                throw "Could not locate the VBE main window for the disposable Excel process."
            }
            $dialog = Wait-ForDialogResult
        }
        finally {
            [XlideVbeDialogWatcher]::Stop()
        }
        if ($null -ne $dialog) {
            Set-DialogOutcome $dialog $true
        }
        else {
            $result.outcome = "accepted"
            $result.stage = $stage
        }
    }
    elseif ($mode -eq "run" -and $case.entryPoint) {
        Set-Stage "run"
        $entryPoint = [string]$case.entryPoint
        $moduleName = [string]$component.Name
        $workbookName = [string]$workbook.Name
        $macroNames = @(
            "'$workbookName'!$moduleName.$entryPoint",
            "'$workbookName'!$entryPoint",
            "$moduleName.$entryPoint",
            $entryPoint
        )
        $lastError = $null
        foreach ($macro in $macroNames) {
            try {
                [void]$excel.Run($macro)
                $lastError = $null
                break
            }
            catch {
                $lastError = $_
                if ($_.Exception.Message -notmatch "Cannot run the macro") {
                    throw
                }
            }
        }
        if ($null -ne $lastError) {
            throw $lastError
        }
    }
    elseif ($mode -ne "run") {
        throw "Unsupported oracle mode: $mode"
    }

    if ($result.outcome -eq "setup_error") {
        $result.outcome = "accepted"
        $result.stage = $stage
    }
}
catch {
    $dialog = Read-DialogResult
    if ($null -ne $dialog) {
        Set-DialogOutcome $dialog ($mode -eq "compile")
    }
    else {
        $result.outcome = "rejected"
        $result.stage = $stage
        $result.message = $_.Exception.Message
        $result.hresult = $_.Exception.HResult
    }
}
finally {
    try { [XlideVbeDialogWatcher]::Stop() } catch { }
    if ($null -ne $workbook) {
        try { $workbook.Close($false) } catch { }
    }
    if ($null -ne $excel) {
        try { $excel.Quit() } catch { }
    }

    Release-ComObject $codeModule
    Release-ComObject $component
    Release-ComObject $vbProject
    Release-ComObject $workbook
    Release-ComObject $excel

    if ($workbookPath -and (Test-Path -LiteralPath $workbookPath)) {
        try { Remove-Item -LiteralPath $workbookPath -Force } catch { }
    }

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

Write-ResultJson $result
