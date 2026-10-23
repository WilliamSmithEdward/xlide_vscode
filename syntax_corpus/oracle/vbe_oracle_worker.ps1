param(
    [Parameter(Mandatory = $false)]
    [string]$CasePath = "",

    [Parameter(Mandatory = $false)]
    [string]$PidPath = "",

    [Parameter(Mandatory = $false)]
    [string]$StagePath = "",

    [Parameter(Mandatory = $false)]
    [string]$DialogPath = "",

    [Parameter(Mandatory = $false)]
    [int]$DialogWatchSeconds = 60,

    [Parameter(Mandatory = $false)]
    [int]$DialogHoldSeconds = 0,

    # Serve cases read from stdin, one JSON request per line, instead of
    # running the single case named by -CasePath.
    [Parameter(Mandatory = $false)]
    [switch]$Serve,

    # How many compile cases one kept host instance serves before it is
    # replaced. -CasesPerExcel is the name it had when only Excel was driven.
    [Parameter(Mandatory = $false)]
    [Alias('CasesPerExcel')]
    [int]$CasesPerInstance = 40,

    # Without -Serve: the host the case runs in, as the coordinator's host
    # table gives it. A served request carries its own.
    [Parameter(Mandatory = $false)]
    [string]$HostName = "excel",

    [Parameter(Mandatory = $false)]
    [string]$HostProgId = "Excel.Application",

    [Parameter(Mandatory = $false)]
    [string]$HostImage = "EXCEL"
)

$ErrorActionPreference = "Stop"

# Information.TypeName reads a designer control's class, as VBA's TypeName does.
Add-Type -AssemblyName Microsoft.VisualBasic

Add-Type @"
using System;
using System.Collections.Generic;
using System.Diagnostics;
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
    private const int IDCANCEL = 2;
    private const int IDABORT = 3;
    private const uint SMTO_ABORTIFHUNG = 0x0002;
    private const int SW_SHOW = 5;
    private const int SW_RESTORE = 9;
    private static Thread watcherThread;
    private static volatile bool stopRequested;
    private static volatile bool dialogSeen;
    private static readonly ManualResetEvent stopEvent = new ManualResetEvent(false);

    public static bool DialogSeen {
        get { return dialogSeen; }
    }

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

    public static void Start(uint hostProcessId, string resultPath, string stagePath, int maxSeconds, int holdSeconds) {
        Start(hostProcessId, resultPath, stagePath, maxSeconds, holdSeconds, false);
    }

    // When a run case starts the watcher with endHostAfterDialog, the watcher
    // ends the host once it has recorded a dialog: the dialog is the case's
    // evidence, and the host is replaced after every run case anyway.
    // Waiting for Run instead can take long. Excel's Run does not return after
    // a compile error is dismissed, so the case would wait for the timeout.
    // In Excel 16.0 (build 20430), after a run-time error in a Function
    // declared As String, Excel crashes in oleaut32 once the dialog is
    // dismissed, and Run returns only when Windows Error Reporting has
    // finished with it, 6 to 10 s later.
    public static void Start(uint hostProcessId, string resultPath, string stagePath, int maxSeconds, int holdSeconds, bool endHostAfterDialog) {
        Stop();
        if (hostProcessId == 0 || String.IsNullOrEmpty(resultPath)) {
            return;
        }

        stopRequested = false;
        dialogSeen = false;
        stopEvent.Reset();
        watcherThread = new Thread(delegate() {
            Watch(hostProcessId, resultPath, stagePath, maxSeconds, holdSeconds, endHostAfterDialog);
        });
        watcherThread.IsBackground = true;
        watcherThread.Start();
    }

    public static void Stop() {
        stopRequested = true;
        stopEvent.Set();
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

    private static void Watch(uint hostProcessId, string resultPath, string stagePath, int maxSeconds, int holdSeconds, bool endHostAfterDialog) {
        DateTime deadline = DateTime.UtcNow.AddSeconds(Math.Max(1, maxSeconds));
        while (!stopRequested && DateTime.UtcNow < deadline) {
            DialogInfo dialog = FindVbeDialog(hostProcessId);
            if (dialog != null) {
                dialogSeen = true;
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
                    if (endHostAfterDialog) {
                        EndProcess(hostProcessId);
                    }
                    return;
                }
            }
            stopEvent.WaitOne(100);
        }
    }

    private static void EndProcess(uint processId) {
        try {
            using (Process process = Process.GetProcessById((int)processId)) {
                process.Kill();
            }
        }
        catch {
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
            stopEvent.WaitOne(100);
        }
    }

    private static DialogInfo FindVbeDialog(uint hostProcessId) {
        DialogInfo found = null;
        EnumWindows(delegate(IntPtr hWnd, IntPtr lParam) {
            if (found != null) {
                return false;
            }

            uint windowProcessId;
            GetWindowThreadProcessId(hWnd, out windowProcessId);
            if (windowProcessId != hostProcessId || !IsWindowVisible(hWnd)) {
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

    private static IntPtr FindVbeMainWindow(uint hostProcessId) {
        IntPtr found = IntPtr.Zero;
        EnumWindows(delegate(IntPtr hWnd, IntPtr lParam) {
            if (found != IntPtr.Zero) {
                return false;
            }

            uint windowProcessId;
            GetWindowThreadProcessId(hWnd, out windowProcessId);
            if (windowProcessId != hostProcessId || !IsWindowVisible(hWnd)) {
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
            SendDialogCommand(hWnd, IDCANCEL, IntPtr.Zero);
            SendDialogCommand(hWnd, IDABORT, IntPtr.Zero);
            SendWindowClose(hWnd);
            DateTime resendAt = DateTime.UtcNow.AddMilliseconds(100);
            while (DateTime.UtcNow < resendAt && DateTime.UtcNow < deadline) {
                if (!IsWindow(hWnd) || !IsWindowVisible(hWnd)) {
                    return true;
                }
                Thread.Sleep(10);
            }
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

# Escaped to ASCII, so a dialog message with any character survives the
# console code page on its way to the coordinator.
function Write-ResultJson($Value) {
    $json = $Value | ConvertTo-Json -Depth 12 -Compress
    return [regex]::Replace($json, '[^\x00-\x7F]', { param($m) '\u{0:x4}' -f [int][char]$m.Value })
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
        return Get-Content -LiteralPath $DialogPath -Raw -Encoding UTF8 | ConvertFrom-Json
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

# No dialog only means accepted if the host is still there to show one. When
# the host crashes, as compiling some snippets makes Excel do, the watcher
# sees no dialog either.
function Assert-HostAlive($App, $HostPid) {
    if ($null -eq (Get-Process -Id $HostPid -ErrorAction SilentlyContinue)) {
        throw "$($script:hostNoun) (pid $HostPid) exited during the $stage stage. A crash is not VBA evidence."
    }
    try {
        $null = $App.Name
    }
    catch {
        throw "$($script:hostNoun) (pid $HostPid) did not answer a COM call after the $stage stage: $($_.Exception.Message)"
    }
}

# A dialog the watcher saw near the end of a wait is recorded by the time
# Stop() has joined it, so look again before concluding there was none.
function Resolve-DialogAfterStop($Dialog) {
    if ($null -ne $Dialog) {
        return $Dialog
    }
    $Dialog = Read-DialogResult
    if ($null -eq $Dialog -and [XlideVbeDialogWatcher]::DialogSeen) {
        throw "A VBE dialog appeared but was not dismissed and recorded."
    }
    return $Dialog
}

# Once the watcher has seen a dialog, its result is on the way: after any
# --dialog-hold-seconds hold and up to 5 s of dismissing. A dialog that was
# seen but never recorded must not count as no dialog.
function Wait-ForSeenDialogResult {
    $deadline = [DateTime]::UtcNow.AddSeconds($DialogHoldSeconds + 10)
    $dialog = Read-DialogResult
    while ($null -eq $dialog -and [DateTime]::UtcNow -lt $deadline) {
        Start-Sleep -Milliseconds 10
        $dialog = Read-DialogResult
    }
    if ($null -eq $dialog) {
        throw "A VBE dialog appeared but was not dismissed and recorded."
    }
    return $dialog
}

# The Compile command is disabled once the project has compiled, and stays
# enabled when compiling fails with an error dialog. A disabled command with no
# dialog is the accept signal, so accepted cases don't wait out the window.
function Wait-ForCompileResult($Control, $HostPid) {
    $deadline = [DateTime]::UtcNow.AddSeconds([Math]::Max(1, $DialogWatchSeconds))
    while ([DateTime]::UtcNow -lt $deadline) {
        if ([XlideVbeDialogWatcher]::DialogSeen) {
            return Wait-ForSeenDialogResult
        }
        # A crashed host answers $null here rather than throwing. A busy one
        # can reject the read, so only a missing process ends the wait early.
        $enabled = $null
        try { $enabled = $Control.Enabled } catch { }
        if ($enabled -isnot [bool]) {
            if ($null -eq (Get-Process -Id $HostPid -ErrorAction SilentlyContinue)) {
                return $null
            }
        }
        elseif (-not $enabled) {
            # Give a dialog that raced the read time to be seen, and make sure
            # the command is still disabled rather than briefly so.
            Start-Sleep -Milliseconds 300
            if ([XlideVbeDialogWatcher]::DialogSeen) {
                return Wait-ForSeenDialogResult
            }
            $enabled = $null
            try { $enabled = $Control.Enabled } catch { }
            if ($enabled -is [bool] -and -not $enabled) {
                return $null
            }
        }
        Start-Sleep -Milliseconds 50
    }
    return $null
}

# A host's Run returns only after a dialog it raised has been dismissed, so
# the watcher has normally seen it by then; 300 ms is a margin, not a wait.
function Wait-ForRunDialogResult {
    $deadline = [DateTime]::UtcNow.AddMilliseconds(300)
    while (-not [XlideVbeDialogWatcher]::DialogSeen -and [DateTime]::UtcNow -lt $deadline) {
        Start-Sleep -Milliseconds 20
    }
    if ([XlideVbeDialogWatcher]::DialogSeen) {
        return Wait-ForSeenDialogResult
    }
    return $null
}

# Errors with which the VBE refuses to take a module's source. They are
# compile evidence, not harness failures. Keep this list to messages that
# have been seen; anything else adding a module fails stays a worker_error.
$InsertionRefusals = @(
    "Too many line continuations"
)

function ConvertTo-NormalizedCommandCaption($Caption) {
    return ([string]$Caption).Replace('&', '').Trim()
}

function ConvertTo-VbComponentType($TypeName) {
    switch (([string]$TypeName).ToLowerInvariant()) {
        "class" { return 2 }
        "userform" { return 3 }       # vbext_ct_MSForm
        "standard" { return 1 }
        default { return 1 }
    }
}

# An Access form or report module, Form_X or Report_X, exists once its form
# or report does and has a module. A new database has none, so the case's
# form or report is created, given a module, saved under its name, and
# closed; the module's code is then replaced like any document module's.
function New-AccessDocumentModule($VbProject, [string]$Name) {
    if ($Name -notmatch '^(Form|Report)_(\w+)$') {
        throw "An Access document module is named Form_<name> or Report_<name>, not '$Name'."
    }
    $isForm = $matches[1] -eq 'Form'
    $objectName = $matches[2]
    $objectType = if ($isForm) { 2 } else { 3 }    # acForm, acReport
    # Assigned directly: through an if expression PowerShell would unroll the
    # form, which enumerates its controls.
    $design = $null
    if ($isForm) { $design = $script:app.CreateForm() } else { $design = $script:app.CreateReport() }
    $design.HasModule = $true
    $created = [string]$design.Name
    Remove-ComObjectReference $design
    $script:app.DoCmd.Close($objectType, $created, 1)    # acSaveYes
    if ($created -ne $objectName) {
        $script:app.DoCmd.Rename($objectName, $objectType, $created)
    }
    return $VbProject.VBComponents.Item($Name)
}

# A new module's code is the case's source and nothing else. Excel keeps what
# its VBE puts in a new module, as it always has; the other hosts drop it, so
# Access's own Option Compare Database is not a second Option statement.
function Set-ComponentSource($Component, [string]$Source) {
    if ($script:hostName -ne "excel") {
        $lineCount = $Component.CodeModule.CountOfLines
        if ($lineCount -gt 0) { $null = $Component.CodeModule.DeleteLines(1, $lineCount) }
    }
    $Component.CodeModule.AddFromString($Source)
}

# What TypeName reports from .NET for a control each ProgID makes: the
# control's default interface, not the class name VBA's TypeName shows.
# Measured in Excel, Word, PowerPoint and Access 16.0 (build 20430); a ProgID
# not listed here is refused until it is measured.
$DesignerControlInterfaces = @{
    "Forms.CommandButton.1" = "ICommandButton"
    "Forms.ComboBox.1" = "IMdcCombo"
}

# A UserForm's controls belong to its designer, not to its code. The case
# names each control and its ProgID; every one is then read back by name and
# type, so a control the designer did not make fails the case as setup, never
# as compile evidence (issue #1419).
function Add-DesignerControls($Component, $ModuleSpec) {
    if (-not ($ModuleSpec.PSObject.Properties.Name -contains "designerControls") -or -not $ModuleSpec.designerControls) {
        return
    }
    $designer = $Component.Designer
    if ($null -eq $designer) {
        throw "Module '$($ModuleSpec.name)' has no designer to take designerControls."
    }
    try {
        foreach ($control in @($ModuleSpec.designerControls)) {
            $added = $designer.Controls.Add([string]$control.progId, [string]$control.name, $true)
            Remove-ComObjectReference $added
        }
        foreach ($control in @($ModuleSpec.designerControls)) {
            $found = $null
            try { $found = $designer.Controls.Item([string]$control.name) } catch { }
            if ($null -eq $found -or [string]$found.Name -cne [string]$control.name) {
                throw "The designer has no control named '$($control.name)' after adding it."
            }
            $expected = $DesignerControlInterfaces[[string]$control.progId]
            $actual = [Microsoft.VisualBasic.Information]::TypeName($found)
            Remove-ComObjectReference $found
            if ($null -eq $expected) {
                throw "No read-back type is known for $($control.progId); add it to `$DesignerControlInterfaces once measured."
            }
            if ($actual -ne $expected) {
                throw "The designer made '$($control.name)' a $actual, not the $expected that $($control.progId) makes."
            }
        }
    }
    finally {
        Remove-ComObjectReference $designer
    }
}

# A document module's code name can be changed, ThisWorkbook to BookCode say.
# The case's `codeName` is set on the module and read back.
function Set-DocumentCodeName($Component, $ModuleSpec) {
    if (-not ($ModuleSpec.PSObject.Properties.Name -contains "codeName") -or -not $ModuleSpec.codeName) {
        return
    }
    $Component.Name = [string]$ModuleSpec.codeName
    if ([string]$Component.Name -cne [string]$ModuleSpec.codeName) {
        throw "Module '$($ModuleSpec.name)' is named '$($Component.Name)', not '$($ModuleSpec.codeName)', after renaming it."
    }
}

# Module attributes (Attribute VB_PredeclaredId = True) are import metadata,
# like procedure attributes. Once imported, the module is exported again and
# every attribute line of the case must be there, or the case fails as setup.
function Assert-ImportedAttributes($Component, [string]$Source, [string]$Extension) {
    $wanted = @([regex]::Matches($Source, '(?m)^[ \t]*(Attribute[ \t]+[^\r\n]*?)[ \t]*\r?$') | ForEach-Object {
        ($_.Groups[1].Value -replace '\s+', ' ').ToLowerInvariant()
    })
    if ($wanted.Count -eq 0) { return }
    $readBack = Join-Path ([IO.Path]::GetTempPath()) ('xlide-oracle-readback-' + [Guid]::NewGuid().ToString('N') + $Extension)
    try {
        $null = $Component.Export($readBack)
        $exported = @([IO.File]::ReadAllLines($readBack, [Text.Encoding]::Default) | ForEach-Object {
            ($_.Trim() -replace '\s+', ' ').ToLowerInvariant()
        })
        foreach ($line in $wanted) {
            if ($exported -notcontains $line) {
                throw "The imported module '$($Component.Name)' lost the attribute line '$line'."
            }
        }
    }
    finally {
        if (Test-Path -LiteralPath $readBack) { Remove-Item -LiteralPath $readBack -Force }
    }
}

function Add-OracleComponent($VbProject, $ModuleSpec) {
    if (([string]$ModuleSpec.type).ToLowerInvariant() -eq "document") {
        # The file's document modules exist already and cannot be added:
        # ThisWorkbook and Sheet1, or ThisDocument, take the case's code in
        # place. The file is disposable, so nothing is put back.
        $component = if ($script:hostName -eq "access") {
            New-AccessDocumentModule $VbProject ([string]$ModuleSpec.name)
        } else {
            $VbProject.VBComponents.Item([string]$ModuleSpec.name)
        }
        Set-DocumentCodeName $component $ModuleSpec
        $lineCount = $component.CodeModule.CountOfLines
        if ($lineCount -gt 0) { $null = $component.CodeModule.DeleteLines(1, $lineCount) }
        $source = [string]$ModuleSpec.source
        if ($source.Trim().Length -gt 0) { $component.CodeModule.AddFromString($source) }
        return $component
    }
    $componentType = ConvertTo-VbComponentType $ModuleSpec.type
    $component = $VbProject.VBComponents.Add($componentType)
    if ($ModuleSpec.PSObject.Properties.Name -contains "name" -and $ModuleSpec.name) {
        $component.Name = [string]$ModuleSpec.name
    }
    Add-DesignerControls $component $ModuleSpec
    $source = [string]$ModuleSpec.source
    # Procedure attributes are import metadata, not editor statements. Adding
    # them as code creates a Syntax error during whole-project compilation.
    # Module attributes, Attribute VB_PredeclaredId = True, are the same
    # (issue #1419); a case that carries them replaces the exported header's.
    $moduleAttributes = $source -match '(?m)^\s*Attribute\s+VB_\w+\s*='
    if ($moduleAttributes -or $source -match '(?m)^\s*Attribute\s+[^\s.]+\.VB_\w+\s*=') {
        if ($componentType -eq 3) {
            throw "A UserForm's attributes come from its designer, not from imported source."
        }
        $extension = if ($componentType -eq 2) { '.cls' } else { '.bas' }
        $importPath = Join-Path ([IO.Path]::GetTempPath()) ('xlide-oracle-component-' + [Guid]::NewGuid().ToString('N') + $extension)
        try {
            $lineCount = $component.CodeModule.CountOfLines
            if ($lineCount -gt 0) { $null = $component.CodeModule.DeleteLines(1, $lineCount) }
            $null = $component.Export($importPath)
            $encoding = [Text.Encoding]::Default
            $header = [IO.File]::ReadAllText($importPath, $encoding)
            if ($moduleAttributes) {
                $redefined = @([regex]::Matches($source, '(?m)^\s*Attribute\s+(VB_\w+)\s*=') | ForEach-Object { $_.Groups[1].Value.ToLowerInvariant() })
                $header = ($header -split '\r?\n' | Where-Object {
                    $_.Trim() -and -not ($_ -match '^\s*Attribute\s+(VB_\w+)\s*=' -and $redefined -contains $matches[1].ToLowerInvariant())
                }) -join "`r`n"
            }
            $null = $VbProject.VBComponents.Remove($component)
            [IO.File]::WriteAllText($importPath, $header + [Environment]::NewLine + $source, $encoding)
            $imported = $VbProject.VBComponents.Import($importPath)
            Assert-ImportedAttributes $imported $source $extension
            return $imported
        } finally {
            if (Test-Path -LiteralPath $importPath) { Remove-Item -LiteralPath $importPath -Force }
        }
    }
    Set-ComponentSource $component $source
    return $component
}

# Where the VBE leaves its selection once its compile dialog is gone: the
# module of the active code pane and the span it selects, 1-based as the VBE
# counts. It is read after the dialog closes, so it is not proof of what the
# dialog highlighted while it was up (issue #1419). Unreadable is $null.
function Get-SelectionAfterDialog($App) {
    $pane = $null
    try {
        $pane = $App.VBE.ActiveCodePane
        if ($null -eq $pane) { return $null }
        $startLine = 0; $startColumn = 0; $endLine = 0; $endColumn = 0
        $pane.GetSelection([ref]$startLine, [ref]$startColumn, [ref]$endLine, [ref]$endColumn)
        return [ordered]@{
            module = [string]$pane.CodeModule.Parent.Name
            startLine = [int]$startLine
            startColumn = [int]$startColumn
            endLine = [int]$endLine
            endColumn = [int]$endColumn
        }
    }
    catch {
        return $null
    }
    finally {
        Remove-ComObjectReference $pane
    }
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

function Invoke-VbeCompileCommand($App) {
    $vbe = $App.VBE
    $projectName = [string]$vbe.ActiveVBProject.Name
    $caption = "Compile $projectName"
    # The Compile command is button (type 1) Id 578. Looking it up directly takes
    # milliseconds; walking every command bar by caption takes about 9 s. The
    # walk stays as the fallback if the lookup fails or finds another control.
    $control = $null
    try {
        $found = $vbe.CommandBars.FindControl(1, 578)
        if ($null -ne $found -and (ConvertTo-NormalizedCommandCaption $found.Caption) -eq $caption) {
            $control = $found
        }
    }
    catch {
    }
    if ($null -eq $control) {
        foreach ($bar in @($vbe.CommandBars)) {
            $control = Find-VbeCommandControl $bar.Controls $caption
            if ($null -ne $control) {
                break
            }
        }
    }
    if ($null -eq $control) {
        throw "Could not find exact VBE command '$caption'. Oracle compile command was not invoked."
    }
    # A disabled command is the accept signal, so it has to start enabled. In
    # a kept instance it can still show the previous project's compiled state.
    $deadline = [DateTime]::UtcNow.AddSeconds(1)
    while ($control.Enabled -ne $true) {
        if ([DateTime]::UtcNow -ge $deadline) {
            throw "The VBE Compile command was disabled before compiling. Oracle compile command was not invoked."
        }
        Start-Sleep -Milliseconds 50
    }
    $control.Execute()
    return $control
}

# One host instance serves several compile cases, each in a fresh file.
# Compiling runs no code, so it leaves the host as it found it. The instance
# is replaced after every run case, whose macro can change the host itself;
# after -CasesPerInstance cases; after any case that did not end in evidence
# or left a file open; when the next case names another host; and when a
# kept instance no longer takes the settings. Access is never kept: closing
# its database asks whether to save the case's modules.
$script:app = $null
$script:appPid = 0
$script:appBuild = $null
$script:appBuildPid = 0
$script:appCases = 0
$script:hostName = $HostName
$script:hostProgId = $HostProgId
$script:hostImage = $HostImage
$script:hostNoun = $HostName
$script:appHost = $null
$script:hostImageOfApp = $null

function Set-OracleHost($Name, $ProgId, $Image) {
    $script:hostName = ([string]$Name).ToLowerInvariant()
    $script:hostProgId = [string]$ProgId
    $script:hostImage = [string]$Image
    $script:hostNoun = switch ($script:hostName) {
        "excel" { "Excel" }
        "word" { "Word" }
        "powerpoint" { "PowerPoint" }
        "access" { "Access" }
        default { throw "Unsupported oracle host: $Name" }
    }
}

function Get-HostProcessIds {
    return @(Get-Process -Name $script:hostImage -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
}

function Get-WindowProcessId([long]$Hwnd) {
    $processId = 0
    [void][XlideUser32]::GetWindowThreadProcessId([IntPtr]$Hwnd, [ref]$processId)
    return [int]$processId
}

function Start-OracleHost {
    $script:appBuild = $null
    $script:appBuildPid = 0
    $before = Get-HostProcessIds
    # Word and PowerPoint can hand an automation request to an instance that
    # is already running, and PowerPoint always does: that instance is
    # someone's, and the oracle would compile into it and then quit it.
    if (($script:hostName -eq "word" -or $script:hostName -eq "powerpoint") -and $before.Count -gt 0) {
        throw "$($script:hostNoun) is already running (pid $($before -join ', ')). The oracle drives only an instance of its own; close $($script:hostNoun) first."
    }
    if ($script:hostName -eq "excel") {
        $script:app = New-Object -ComObject Excel.Application
    }
    else {
        # Late-bound, so Word's Quit(0) and Access's Quit(2) take plain values.
        $script:app = [Activator]::CreateInstance([type]::GetTypeFromProgID($script:hostProgId))
    }
    $processId = switch ($script:hostName) {
        "excel" { Get-WindowProcessId $script:app.Hwnd }
        "powerpoint" { Get-WindowProcessId $script:app.HWND }
        "access" { Get-WindowProcessId $script:app.hWndAccessApp() }
        default {
            # Word's Application has no window handle of its own: the instance
            # is the process that was not there before.
            $new = @(Get-HostProcessIds | Where-Object { $before -notcontains $_ })
            if ($new.Count -eq 1) { $new[0] } else { 0 }
        }
    }
    if (-not $processId) {
        $script:app = $null
        throw "Could not tell which $($script:hostNoun) process the oracle started."
    }
    $script:appPid = $processId
    $script:appHost = $script:hostName
    $script:appCases = 0
}

# The host this worker drives, for the coordinator: its process id, image and
# start time (issue #1439). The coordinator ends a host only while the process
# at that id still has that image and start time, so a reused id is never
# taken for it. Written as soon as the host is started, and before each case.
function Write-HostRecord {
    $start = (Get-Process -Id $script:appPid -ErrorAction Stop).StartTime.ToUniversalTime().ToFileTimeUtc()
    Set-Content -LiteralPath $PidPath -Value "$($script:appPid) $($script:hostImageOfApp) $start" -Encoding ascii
}

function Stop-OracleHost {
    $script:appBuild = $null
    $script:appBuildPid = 0
    if ($null -ne $script:app) {
        try {
            switch ($script:appHost) {
                "word" { $script:app.Quit(0) }
                "access" { $script:app.Quit(2) }
                default { $script:app.Quit() }
            }
        } catch { }
        Remove-ComObjectReference $script:app
        $script:app = $null
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
    if ($script:appPid) {
        # End only the process this worker recorded, and only if it is still
        # the host's: a pid can have been reused.
        Get-Process -Id $script:appPid -ErrorAction SilentlyContinue |
            Where-Object { $_.ProcessName -eq $script:hostImageOfApp } |
            Stop-Process -Force -ErrorAction SilentlyContinue
        # Gone before the next case looks, or the next Word or PowerPoint case
        # would take this worker's own exiting instance for someone else's.
        $deadline = [DateTime]::UtcNow.AddSeconds(10)
        while ([DateTime]::UtcNow -lt $deadline -and $null -ne (Get-Process -Id $script:appPid -ErrorAction SilentlyContinue)) {
            Start-Sleep -Milliseconds 50
        }
        # A host still running is not forgotten (issue #1439): the worker fails,
        # and the coordinator, which recorded it, ends it or reports it.
        $left = Get-Process -Id $script:appPid -ErrorAction SilentlyContinue | Where-Object { $_.ProcessName -eq $script:hostImageOfApp }
        if ($null -ne $left) {
            throw "The $($script:hostNoun) instance the oracle started, pid $($script:appPid), did not exit."
        }
        $script:appPid = 0
    }
    $script:appHost = $null
}

# Applied before every case, since a case's macro can change them.
function Set-OracleHostSettings($App) {
    # The host must stay visible: when it is hidden, the VBE Compile-error /
    # Run-time-error modal is not surfaced as a detectable visible top-level
    # window, so the Win32 dialog watcher never sees it and a rejection is
    # silently misrecorded as "accepted". The worker still does not require
    # foreground focus. AutomationSecurity 1 is msoAutomationSecurityLow,
    # for the disposable instance this worker created.
    switch ($script:hostName) {
        "excel" {
            $App.Visible = $true
            $App.DisplayAlerts = $false
            $App.EnableEvents = $false
            try { $App.ScreenUpdating = $false } catch { }
        }
        "word" {
            $App.Visible = $true
            $App.DisplayAlerts = 0            # wdAlertsNone
        }
        "powerpoint" {
            $App.Visible = -1                 # msoTrue
            $App.DisplayAlerts = 1            # ppAlertsNone
        }
        "access" {
            $App.Visible = $true
            # Refused until a database is open, so it is asked again after.
            try { $App.DoCmd.SetWarnings($false) } catch { }
        }
    }
    try { $App.AutomationSecurity = 1 } catch { }
}

# A fresh file for one case. Access's database has to be on disk; its path
# comes back so the case can delete it.
function New-OracleContainer($App) {
    switch ($script:hostName) {
        "excel" { $container = $App.Workbooks.Add(); [void]$container.Activate(); return @{ Container = $container; Path = $null } }
        "word" { return @{ Container = $App.Documents.Add(); Path = $null } }
        "powerpoint" { return @{ Container = $App.Presentations.Add(-1); Path = $null } }
        "access" {
            # A plain string: given Join-Path's wrapped one, NewCurrentDatabase
            # never returns (measured 2026-10-16; the same call from Python
            # takes 0.1 s).
            $path = [string](Join-Path ([IO.Path]::GetTempPath()) ('XlideOracle_' + [Guid]::NewGuid().ToString('N') + '.accdb'))
            $App.NewCurrentDatabase($path)
            $App.DoCmd.SetWarnings($false)
            return @{ Container = $null; Path = $path }
        }
    }
}

function Get-OracleVbProject($App, $Container) {
    if ($script:hostName -eq "access") {
        return $App.VBE.ActiveVBProject
    }
    return $Container.VBProject
}

function Close-OracleContainer($App, $Container) {
    switch ($script:hostName) {
        "excel" { $Container.Close($false) }
        "word" { $Container.Close(0) }    # wdDoNotSaveChanges
        "powerpoint" { $Container.Saved = -1; $Container.Close() }
        "access" { }                        # the instance quits without saving
    }
}

function Get-OpenContainerCount($App) {
    switch ($script:hostName) {
        "excel" { return $App.Workbooks.Count }
        "word" { return $App.Documents.Count }
        "powerpoint" { return $App.Presentations.Count }
        default { return 0 }
    }
}

# The names a host's Run might know the entry point by, most specific first,
# and how it says it found none.
function Get-RunMacroNames($Container, $Component, [string]$EntryPoint, $VbProject) {
    $moduleName = [string]$Component.Name
    switch ($script:hostName) {
        "excel" {
            $fileName = [string]$Container.Name
            return @("'$fileName'!$moduleName.$EntryPoint", "'$fileName'!$EntryPoint", "$moduleName.$EntryPoint", $EntryPoint)
        }
        "word" {
            $projectName = [string]$VbProject.Name
            return @("$projectName.$moduleName.$EntryPoint", "$moduleName.$EntryPoint", $EntryPoint)
        }
        "powerpoint" {
            $fileName = [string]$Container.Name
            return @("'$fileName'!$moduleName.$EntryPoint", "$moduleName.$EntryPoint", $EntryPoint)
        }
        "access" { return @($EntryPoint) }
    }
}

$MacroNotFoundPatterns = @(
    "Cannot run the macro",
    "Unable to run the specified macro",
    "Can't run the specified macro",
    "Macro cannot be found",
    "cannot find the procedure"
)

function Test-MacroNotFound([string]$Message) {
    foreach ($pattern in $MacroNotFoundPatterns) {
        if ($Message.Contains($pattern)) { return $true }
    }
    return $false
}

# The build of the running host, read from its own executable once per host
# process. A failed read is not kept, so the next case reads it again.
function Get-HostBuild($HostPid) {
    if ($null -ne $script:appBuild -and $script:appBuildPid -eq $HostPid) {
        return $script:appBuild
    }
    $script:appBuild = $null
    $script:appBuildPid = 0
    try {
        $path = (Get-Process -Id $HostPid -ErrorAction Stop).MainModule.FileName
        $v = [Diagnostics.FileVersionInfo]::GetVersionInfo($path)
        if ($v.FileMajorPart -gt 0) {
            $script:appBuild = '{0}.{1}.{2}.{3}' -f $v.FileMajorPart, $v.FileMinorPart, $v.FileBuildPart, $v.FilePrivatePart
            $script:appBuildPid = $HostPid
            return $script:appBuild
        }
    } catch { }
    return $null
}

function Invoke-OracleCase {
    # The coordinator writes the case as UTF-8; Windows PowerShell would
    # otherwise read it in the ANSI code page and garble non-ASCII source.
    $case = Get-Content -LiteralPath $CasePath -Raw -Encoding UTF8 | ConvertFrom-Json
    $app = $null
    $appPid = 0
    $container = $null
    $vbProject = $null
    $component = $null
    $codeModule = $null
    $containerPath = $null
    $mode = "compile"
    Set-Stage "setup"

    $result = [ordered]@{
        caseId = $case.id
        outcome = "setup_error"
        stage = $stage
        message = ""
        hresult = $null
        host = $script:hostName
        hostBuild = $null
    }

    try {
        Set-Stage "start_host"
        # A kept instance of another host, or one that no longer takes the
        # settings, is replaced before the case starts rather than failing it.
        if ($null -ne $script:app -and $script:appHost -ne $script:hostName) {
            Stop-OracleHost
        }
        if ($null -ne $script:app) {
            Write-HostRecord
            try { Set-OracleHostSettings $script:app } catch { Stop-OracleHost }
        }
        if ($null -eq $script:app) {
            $script:hostImageOfApp = $script:hostImage
            Start-OracleHost
            Write-HostRecord
            Set-OracleHostSettings $script:app
        }
        $app = $script:app
        $appPid = $script:appPid
        # The build of the running host, not of the registration. Failure leaves
        # it unknown; checkpoint runs reject such evidence.
        $result.hostBuild = Get-HostBuild $appPid
        $script:appCases += 1

        Set-Stage "create_file"
        $created = New-OracleContainer $app
        $container = $created.Container
        $containerPath = $created.Path

        Set-Stage "add_module"
        $vbProject = Get-OracleVbProject $app $container
        $component = $null
        # Type libraries the case's code names early bound, added by GUID and
        # version before any module, as Tools > References would.
        if ($case.PSObject.Properties.Name -contains "references" -and $case.references) {
            Set-Stage "add_references"
            foreach ($reference in @($case.references)) {
                $null = $vbProject.References.AddFromGuid([string]$reference.guid, [int]$reference.major, [int]$reference.minor)
            }
        }
        if ($case.PSObject.Properties.Name -contains "modules" -and $case.modules) {
            Set-Stage "add_modules"
            foreach ($moduleSpec in @($case.modules)) {
                $added = Add-OracleComponent $vbProject $moduleSpec
                if ($null -eq $component -and (ConvertTo-VbComponentType $moduleSpec.type) -eq 1) {
                    $component = $added
                }
                if ($moduleSpec.PSObject.Properties.Name -contains "entry" -and $moduleSpec.entry) {
                    $component = $added
                }
            }
            if ($null -eq $component) {
                throw "Oracle case with modules must include a standard entry module."
            }
        }
        else {
            $component = $vbProject.VBComponents.Add(1)
            $component.Name = "XlideOracleModule"
            Set-ComponentSource $component ([string]$case.source)
        }
        $codeModule = $component.CodeModule

        if ($case.PSObject.Properties.Name -contains "mode" -and $case.mode) {
            $mode = [string]$case.mode
        }

        if ($mode -eq "compile") {
            Set-Stage "compile"
            # Compile probes drive the VBE command bars; the VBE main window is shown
            # so the Compile command is available and its error dialog is surfaced.
            $app.VBE.MainWindow.Visible = $true
            # The Compile command compiles the active project, and every case's
            # project has the same name, so make sure it is this case's.
            $app.VBE.ActiveVBProject = $vbProject

            [void]$component.Activate()
            $codeModule.CodePane.Show()
            $codeModule.CodePane.SetSelection(1, 1, 1, 1)

            [XlideVbeDialogWatcher]::Start($appPid, $DialogPath, $StagePath, $DialogWatchSeconds, $DialogHoldSeconds)
            try {
                $compileControl = Invoke-VbeCompileCommand $app
                $dialog = Wait-ForCompileResult $compileControl $appPid
            }
            finally {
                [XlideVbeDialogWatcher]::Stop()
            }
            $dialog = Resolve-DialogAfterStop $dialog
            if ($null -ne $dialog) {
                Set-DialogOutcome $dialog $true
                $selection = Get-SelectionAfterDialog $app
                if ($null -ne $selection) { $result.selectionAfterDialog = $selection }
            }
            else {
                Assert-HostAlive $app $appPid
                $result.outcome = "accepted"
                $result.stage = $stage
            }
        }
        elseif ($mode -eq "run" -and $case.entryPoint) {
            Set-Stage "run"
            $macroNames = Get-RunMacroNames $container $component ([string]$case.entryPoint) $vbProject
            $lastError = $null
            [XlideVbeDialogWatcher]::Start($appPid, $DialogPath, $StagePath, $DialogWatchSeconds, $DialogHoldSeconds, $true)
            try {
                foreach ($macro in $macroNames) {
                    try {
                        [void]$app.Run($macro)
                        $lastError = $null
                        break
                    }
                    catch {
                        # The macro ran, but its returned Date lies beyond what
                        # .NET can hold, so only reading the return value failed.
                        # This one exact error counts as a return; any other
                        # stays an error.
                        $returnError = $_.Exception.InnerException
                        if ($_.Exception -is [System.Management.Automation.MethodInvocationException] -and
                            $null -ne $returnError -and
                            $returnError.GetType() -eq [System.ArgumentException] -and
                            $returnError.Message -ceq "Not a legal OleAut date.") {
                            $lastError = $null
                            break
                        }
                        $lastError = $_
                        if (-not (Test-MacroNotFound $_.Exception.Message)) {
                            throw
                        }
                    }
                }
                if ($null -ne $lastError) {
                    throw $lastError
                }
            }
            finally {
                $dialog = Wait-ForRunDialogResult
                [XlideVbeDialogWatcher]::Stop()
            }
            $dialog = Resolve-DialogAfterStop $dialog
            if ($null -ne $dialog) {
                Set-DialogOutcome $dialog $false
            }
        }

        elseif ($mode -ne "run") {
            throw "Unsupported oracle mode: $mode"
        }

        if ($result.outcome -eq "setup_error") {
            Assert-HostAlive $app $appPid
            $result.outcome = "accepted"
            $result.stage = $stage
        }
    }
    catch {
        $dialog = Read-DialogResult
        if ($null -ne $dialog) {
            Set-DialogOutcome $dialog ($mode -eq "compile")
        }
        elseif (($stage -eq "add_module" -or $stage -eq "add_modules") -and $InsertionRefusals -contains $_.Exception.Message) {
            # The VBE refused the source itself, as its editor would.
            $result.outcome = "rejected"
            $result.stage = "insert_refused"
            $result.evidencePhase = "compile"
            $result.message = $_.Exception.Message
            $result.hresult = $_.Exception.HResult
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
        $closed = $true
        if ($null -ne $container) {
            try { Close-OracleContainer $app $container } catch { $closed = $false }
        }

        Remove-ComObjectReference $codeModule
        Remove-ComObjectReference $component
        Remove-ComObjectReference $vbProject
        Remove-ComObjectReference $container

        # A macro can open files of its own; they must not outlive the case.
        if ($closed -and $null -ne $script:app) {
            try { $closed = (Get-OpenContainerCount $script:app) -eq 0 } catch { $closed = $false }
        }
        $evidence = $result.outcome -eq "accepted" -or $result.outcome -eq "rejected"
        if (-not $closed -or -not $evidence -or $mode -ne "compile" -or $script:hostName -eq "access" -or $script:appCases -ge $CasesPerInstance) {
            Stop-OracleHost
        }

        # Access's database is a file of its own, free once Access has quit.
        if ($containerPath) {
            foreach ($candidate in @($containerPath, [IO.Path]::ChangeExtension($containerPath, '.laccdb'))) {
                if (Test-Path -LiteralPath $candidate) {
                    try { Remove-Item -LiteralPath $candidate -Force } catch { }
                }
            }
        }
    }

    return $result
}

if ($Serve) {
    try {
        while ($null -ne ($line = [Console]::In.ReadLine())) {
            if (-not $line.Trim()) {
                continue
            }
            $request = $line | ConvertFrom-Json
            $script:CasePath = [string]$request.casePath
            $script:PidPath = [string]$request.pidPath
            $script:StagePath = [string]$request.stagePath
            $script:DialogPath = [string]$request.dialogPath
            $script:DialogWatchSeconds = [int]$request.dialogWatchSeconds
            $script:DialogHoldSeconds = [int]$request.dialogHoldSeconds
            Set-OracleHost $request.host $request.hostProgId $request.hostImage
            $output = @(Invoke-OracleCase)
            [Console]::Out.WriteLine("XLIDE-ORACLE-RESULT " + (Write-ResultJson $output[-1]))
            [Console]::Out.Flush()
        }
    }
    finally {
        Stop-OracleHost
    }
}
else {
    if (-not $CasePath -or -not $PidPath) {
        throw "-CasePath and -PidPath are required without -Serve."
    }
    Set-OracleHost $HostName $HostProgId $HostImage
    $output = @(Invoke-OracleCase)
    Stop-OracleHost
    Write-ResultJson $output[-1]
}
