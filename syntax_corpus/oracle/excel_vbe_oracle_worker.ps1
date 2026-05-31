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
    [int]$DialogWatchSeconds = 60,

    [Parameter(Mandatory = $false)]
    [int]$DialogHoldSeconds = 0
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
    private const int BM_CLICK = 0x00F5;
    private const int IDOK = 1;
    private const uint SMTO_ABORTIFHUNG = 0x0002;
    private const int SW_SHOW = 5;
    private const int SW_RESTORE = 9;
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
    private static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern IntPtr SendMessageTimeout(
        IntPtr hWnd,
        uint Msg,
        IntPtr wParam,
        IntPtr lParam,
        uint fuFlags,
        uint uTimeout,
        out IntPtr lpdwResult
    );

    [DllImport("user32.dll")]
    private static extern int GetDlgCtrlID(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern bool SetForegroundWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

    [DllImport("user32.dll")]
    private static extern bool IsIconic(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern IntPtr GetForegroundWindow();

    [DllImport("user32.dll")]
    private static extern bool AttachThreadInput(uint idAttach, uint idAttachTo, bool fAttach);

    [DllImport("user32.dll")]
    private static extern bool BringWindowToTop(IntPtr hWnd);

    [DllImport("kernel32.dll")]
    private static extern uint GetCurrentThreadId();

    public static void Start(uint excelProcessId, string resultPath, string stagePath, int maxSeconds, int holdSeconds) {
        Stop();
        if (excelProcessId == 0 || String.IsNullOrEmpty(resultPath)) {
            return;
        }

        stopRequested = false;
        watcherThread = new Thread(delegate() {
            Watch(excelProcessId, resultPath, stagePath, maxSeconds, holdSeconds);
        });
        watcherThread.IsBackground = true;
        watcherThread.Start();
    }

    public static void Stop() {
        stopRequested = true;
        if (watcherThread != null && watcherThread.IsAlive) {
            watcherThread.Join(5000);
        }
        watcherThread = null;
    }

    private static bool FocusWindow(IntPtr hWnd) {
        if (hWnd == IntPtr.Zero || !IsWindow(hWnd)) {
            return false;
        }

        if (IsIconic(hWnd)) {
            ShowWindow(hWnd, SW_RESTORE);
        }
        else {
            ShowWindow(hWnd, SW_SHOW);
        }

        IntPtr foreground = GetForegroundWindow();
        uint foregroundProcessId;
        uint foregroundThreadId = foreground == IntPtr.Zero
            ? 0
            : GetWindowThreadProcessId(foreground, out foregroundProcessId);
        uint targetProcessId;
        uint targetThreadId = GetWindowThreadProcessId(hWnd, out targetProcessId);
        uint currentThreadId = GetCurrentThreadId();
        bool attachedForeground = false;
        bool attachedTarget = false;
        try {
            if (foregroundThreadId != 0 && foregroundThreadId != currentThreadId) {
                attachedForeground = AttachThreadInput(currentThreadId, foregroundThreadId, true);
            }
            if (targetThreadId != 0 && targetThreadId != currentThreadId) {
                attachedTarget = AttachThreadInput(currentThreadId, targetThreadId, true);
            }
            BringWindowToTop(hWnd);
            SetForegroundWindow(hWnd);
        }
        finally {
            if (attachedTarget) {
                AttachThreadInput(currentThreadId, targetThreadId, false);
            }
            if (attachedForeground) {
                AttachThreadInput(currentThreadId, foregroundThreadId, false);
            }
        }

        Thread.Sleep(250);
        return GetForegroundWindow() == hWnd;
    }

    private static void Watch(uint excelProcessId, string resultPath, string stagePath, int maxSeconds, int holdSeconds) {
        DateTime deadline = DateTime.UtcNow.AddSeconds(Math.Max(1, maxSeconds));
        while (!stopRequested && DateTime.UtcNow < deadline) {
            DialogInfo dialog = FindVbeDialog(excelProcessId);
            if (dialog != null) {
                if (!String.IsNullOrEmpty(stagePath)) {
                    try {
                        string stage = dialog.Kind == "vbe_compile_dialog" ? "compile_dialog" : "vbe_dialog";
                        File.WriteAllText(stagePath, stage, Encoding.ASCII);
                    }
                    catch {
                    }
                }
                HoldDialogForInspection(dialog.Handle, holdSeconds);
                if (DismissDialog(dialog.Handle)) {
                    WriteDialogResult(resultPath, dialog);
                    return;
                }
            }
            Thread.Sleep(100);
        }
    }

    private static void HoldDialogForInspection(IntPtr hWnd, int holdSeconds) {
        int seconds = Math.Max(0, holdSeconds);
        if (seconds <= 0) {
            return;
        }
        FocusWindow(hWnd);
        DateTime deadline = DateTime.UtcNow.AddSeconds(seconds);
        while (!stopRequested && DateTime.UtcNow < deadline && IsWindow(hWnd)) {
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
            if (className != "#32770" || !IsVbeDialogTitle(title)) {
                return true;
            }

            List<string> childTexts = GetChildTexts(hWnd);
            if (!LooksLikeVbeDialog(childTexts)) {
                return true;
            }
            found = new DialogInfo(hWnd, title, className, childTexts);
            return false;
        }, IntPtr.Zero);
        return found;
    }

    private static bool IsVbeDialogTitle(string title) {
        return title == "Microsoft Visual Basic for Applications" || title == "Microsoft Visual Basic";
    }

    private static bool LooksLikeVbeDialog(List<string> texts) {
        for (int i = 0; i < texts.Count; i++) {
            string text = texts[i];
            if (text.Contains("Compile error:") || text.Contains("Run-time error")) {
                return true;
            }
        }
        return false;
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

    private static bool DismissDialog(IntPtr hWnd) {
        bool closed = DismissDialog(hWnd, 5000);
        return closed;
    }

    private static bool DismissDialog(IntPtr hWnd, int maxMilliseconds) {
        DateTime deadline = DateTime.UtcNow.AddMilliseconds(Math.Max(1, maxMilliseconds));
        while (DateTime.UtcNow < deadline) {
            if (!IsWindow(hWnd) || !IsWindowVisible(hWnd)) {
                return true;
            }

            IntPtr button = FindDialogButton(hWnd, "OK", "End");
            if (button != IntPtr.Zero) {
                int controlId = GetDlgCtrlID(button);
                if (controlId == 0) {
                    controlId = IDOK;
                }
                SendDialogCommand(hWnd, controlId, button);
                SendButtonClick(button);
            }
            SendDialogCommand(hWnd, IDOK, IntPtr.Zero);
            SendWindowClose(hWnd);
            Thread.Sleep(100);
        }
        return !IsWindow(hWnd) || !IsWindowVisible(hWnd);
    }

    private static IntPtr FindDialogButton(IntPtr parent, params string[] captions) {
        IntPtr found = IntPtr.Zero;
        EnumChildWindows(parent, delegate(IntPtr hWnd, IntPtr lParam) {
            string text = GetText(hWnd).Trim().Replace("&", "");
            for (int i = 0; i < captions.Length; i++) {
                if (String.Equals(text, captions[i], StringComparison.OrdinalIgnoreCase)) {
                    found = hWnd;
                    return false;
                }
            }
            return true;
        }, IntPtr.Zero);
        return found;
    }

    private static void SendDialogCommand(IntPtr parent, int controlId, IntPtr control) {
        IntPtr ignored;
        SendMessageTimeout(
            parent,
            WM_COMMAND,
            new IntPtr(controlId),
            control,
            SMTO_ABORTIFHUNG,
            500,
            out ignored
        );
        PostMessage(parent, WM_COMMAND, new IntPtr(controlId), control);
    }

    private static void SendButtonClick(IntPtr button) {
        IntPtr ignored;
        SendMessageTimeout(
            button,
            BM_CLICK,
            IntPtr.Zero,
            IntPtr.Zero,
            SMTO_ABORTIFHUNG,
            500,
            out ignored
        );
    }

    private static void SendWindowClose(IntPtr hWnd) {
        IntPtr ignored;
        SendMessageTimeout(
            hWnd,
            WM_CLOSE,
            IntPtr.Zero,
            IntPtr.Zero,
            SMTO_ABORTIFHUNG,
            500,
            out ignored
        );
        PostMessage(hWnd, WM_CLOSE, IntPtr.Zero, IntPtr.Zero);
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
                if (texts[i].Contains("Run-time error")) {
                    return "vbe_runtime_dialog";
                }
            }
            return "vbe_dialog";
        }

        private static string BuildMessage(List<string> texts) {
            List<string> messageParts = new List<string>();
            for (int i = 0; i < texts.Count; i++) {
                string text = texts[i];
                string normalized = text.Replace("&", "");
                if (normalized == "OK" || normalized == "Help" || normalized == "Continue" || normalized == "End" || normalized == "Debug") {
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

function Remove-ComObjectReference($Value) {
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
        $result.stage = "runtime_dialog"
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

function ConvertTo-NormalizedCommandCaption($Caption) {
    return ([string]$Caption).Replace('&', '').Trim()
}

function Find-VbeCommandControl($Controls, [string]$Caption) {
    foreach ($control in @($Controls)) {
        $normalized = ConvertTo-NormalizedCommandCaption $control.Caption
        if ($normalized -eq $Caption) {
            return $control
        }
        try {
            if ($control.Controls -and $control.Controls.Count -gt 0) {
                $nested = Find-VbeCommandControl $control.Controls $Caption
                if ($null -ne $nested) {
                    return $nested
                }
            }
        }
        catch {
        }
    }
    return $null
}

function Invoke-VbeCompileCommand($Excel) {
    $vbe = $Excel.VBE
    $projectName = [string]$vbe.ActiveVBProject.Name
    $caption = "Compile $projectName"
    foreach ($bar in @($vbe.CommandBars)) {
        $control = Find-VbeCommandControl $bar.Controls $caption
        if ($null -ne $control) {
            $control.Execute()
            return
        }
    }
    throw "Could not find exact VBE command '$caption'. Oracle compile command was not invoked."
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

        [XlideVbeDialogWatcher]::Start($excelPid, $DialogPath, $StagePath, $DialogWatchSeconds, $DialogHoldSeconds)
        try {
            Invoke-VbeCompileCommand $excel
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
        [XlideVbeDialogWatcher]::Start($excelPid, $DialogPath, $StagePath, $DialogWatchSeconds, $DialogHoldSeconds)
        try {
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
        finally {
            $dialog = Wait-ForDialogResult
            [XlideVbeDialogWatcher]::Stop()
        }
        if ($null -ne $dialog) {
            Set-DialogOutcome $dialog $false
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
        $result.outcome = "worker_error"
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

    Remove-ComObjectReference $codeModule
    Remove-ComObjectReference $component
    Remove-ComObjectReference $vbProject
    Remove-ComObjectReference $workbook
    Remove-ComObjectReference $excel

    if ($workbookPath -and (Test-Path -LiteralPath $workbookPath)) {
        try { Remove-Item -LiteralPath $workbookPath -Force } catch { }
    }

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

Write-ResultJson $result
