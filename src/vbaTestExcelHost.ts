import type { VbaTestCase } from './vbaTestRunner';
import type { VbaTestHostOracleEvent } from './vbaTestHostOracle';

export const XLIDE_TEST_HOST_EVENT_PREFIX = 'XLIDE_TEST_HOST_EVENT|';
export const DEFAULT_VBA_TEST_TIMEOUT_MS = 30000;

export interface VbaTestHostPlanItem {
    qualifiedName: string;
    timeoutMs: number;
    expectedFailure: boolean;
}

export interface OwnedReadOnlyExcelTestHostScriptOptions {
    failFast?: boolean;
}

export function vbaTestHostPlanItems(tests: readonly VbaTestCase[]): VbaTestHostPlanItem[] {
    return tests.map((test) => ({
        qualifiedName: test.qualifiedName,
        timeoutMs: test.metadata.timeoutMs ?? DEFAULT_VBA_TEST_TIMEOUT_MS,
        expectedFailure: Boolean(test.metadata.xfailReason),
    }));
}

function productionModalWatcherCSharp(): string {
    return String.raw`
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;

public static class XlideTestModalWatcher
{
    private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern bool EnumChildWindows(IntPtr hWndParent, EnumWindowsProc lpEnumFunc, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern bool IsWindowVisible(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

    [DllImport("user32.dll")]
    private static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

    [DllImport("user32.dll")]
    private static extern int GetWindowTextLength(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern IntPtr SendMessageTimeout(
        IntPtr hWnd,
        uint Msg,
        IntPtr wParam,
        IntPtr lParam,
        uint fuFlags,
        uint uTimeout,
        out IntPtr lpdwResult);

    private const uint BM_CLICK = 0x00F5;
    private const uint SMTO_ABORTIFHUNG = 0x0002;

    private static readonly object Gate = new object();
    private static Timer WatcherTimer;
    private static uint ExcelPid;
    private static string EventPrefix = "";
    private static string ExcelId = "";
    private static string QualifiedName = "";
    private static HashSet<string> Seen = new HashSet<string>(StringComparer.Ordinal);

    private sealed class DialogInfo
    {
        public IntPtr Handle;
        public string ClassName = "";
        public string Title = "";
        public string Message = "";
        public readonly List<string> Texts = new List<string>();
        public readonly List<ButtonInfo> Buttons = new List<ButtonInfo>();
    }

    private sealed class ButtonInfo
    {
        public IntPtr Handle;
        public string Text = "";
    }

    private sealed class DialogAction
    {
        public bool SafeToDismiss;
        public ButtonInfo Button;
        public string Classification = "excel-modal";
        public string Reason = "unsafe-or-unknown-dialog";
    }

    public static void Start(uint processId, string eventPrefix, string excelId, string qualifiedName)
    {
        Stop();
        if (processId == 0)
        {
            return;
        }

        lock (Gate)
        {
            ExcelPid = processId;
            EventPrefix = eventPrefix ?? "";
            ExcelId = excelId ?? "";
            QualifiedName = qualifiedName ?? "";
            Seen = new HashSet<string>(StringComparer.Ordinal);
            WatcherTimer = new Timer(_ => Scan(), null, 200, 250);
        }
    }

    public static void Stop()
    {
        lock (Gate)
        {
            if (WatcherTimer != null)
            {
                WatcherTimer.Dispose();
                WatcherTimer = null;
            }
            ExcelPid = 0;
        }
    }

    private static void Scan()
    {
        uint processId;
        lock (Gate)
        {
            processId = ExcelPid;
        }

        if (processId == 0)
        {
            return;
        }

        try
        {
            EnumWindows((hWnd, lParam) =>
            {
                try
                {
                    InspectTopLevelWindow(hWnd, processId);
                }
                catch
                {
                }
                return true;
            }, IntPtr.Zero);
        }
        catch
        {
        }
    }

    private static void InspectTopLevelWindow(IntPtr hWnd, uint processId)
    {
        if (!IsWindowVisible(hWnd))
        {
            return;
        }

        uint ownerPid;
        GetWindowThreadProcessId(hWnd, out ownerPid);
        if (ownerPid != processId)
        {
            return;
        }

        string className = WindowClass(hWnd);
        if (!String.Equals(className, "#32770", StringComparison.Ordinal))
        {
            return;
        }

        DialogInfo info = BuildDialogInfo(hWnd, className);
        DialogAction action = ChooseAction(info);
        string key = DialogKey(info);
        bool shouldEmitDetected;
        lock (Gate)
        {
            shouldEmitDetected = Seen.Add("detected|" + key);
        }
        if (shouldEmitDetected)
        {
            EmitDetected(info, action);
        }

        if (!action.SafeToDismiss || action.Button == null)
        {
            EmitBlockedOnce(info, action, key);
            return;
        }

        bool dismissed = ClickButton(action.Button.Handle);
        bool shouldEmitDismissed;
        lock (Gate)
        {
            shouldEmitDismissed = Seen.Add("dismissed|" + key);
        }
        if (shouldEmitDismissed)
        {
            EmitDismissed(info, action.Button.Text, dismissed);
        }
        if (!dismissed)
        {
            action.SafeToDismiss = false;
            action.Reason = "low-level-dismiss-failed";
            EmitBlockedOnce(info, action, key);
        }
    }

    private static DialogInfo BuildDialogInfo(IntPtr hWnd, string className)
    {
        DialogInfo info = new DialogInfo();
        info.Handle = hWnd;
        info.ClassName = className;
        info.Title = WindowText(hWnd).Trim();
        EnumChildWindows(hWnd, (child, lParam) =>
        {
            string text = WindowText(child).Trim();
            if (text.Length == 0)
            {
                return true;
            }

            string childClass = WindowClass(child);
            if (String.Equals(childClass, "Button", StringComparison.OrdinalIgnoreCase))
            {
                info.Buttons.Add(new ButtonInfo { Handle = child, Text = text });
            }
            else if (!info.Texts.Contains(text))
            {
                info.Texts.Add(text);
            }
            return true;
        }, IntPtr.Zero);
        info.Message = PickMessage(info);
        return info;
    }

    private static DialogAction ChooseAction(DialogInfo info)
    {
        DialogAction action = new DialogAction();
        action.Classification = Classify(info);

        ButtonInfo endButton = FindButton(info, "End");
        if (endButton != null &&
            (String.Equals(action.Classification, "runtime-error", StringComparison.Ordinal) ||
             String.Equals(action.Classification, "compile-error", StringComparison.Ordinal) ||
             String.Equals(action.Classification, "vba-modal", StringComparison.Ordinal)))
        {
            action.SafeToDismiss = true;
            action.Button = endButton;
            action.Reason = "dismiss-vbe-error-end";
            return action;
        }

        ButtonInfo okButton = FindButton(info, "OK");
        if (okButton != null && HasOnlyInformationalButtons(info))
        {
            action.SafeToDismiss = true;
            action.Button = okButton;
            action.Reason = "dismiss-informational-ok";
            return action;
        }

        action.SafeToDismiss = false;
        action.Reason = "decision-or-unknown-dialog";
        return action;
    }

    private static string Classify(DialogInfo info)
    {
        string haystack = (info.Title + " " + info.Message + " " + Join(info.Texts, " ")).ToLowerInvariant();
        if (haystack.Contains("compile error"))
        {
            return "compile-error";
        }
        if (haystack.Contains("run-time error") || haystack.Contains("runtime error"))
        {
            return "runtime-error";
        }
        if (haystack.Contains("microsoft visual basic"))
        {
            return "vba-modal";
        }
        return "excel-modal";
    }

    private static bool HasOnlyInformationalButtons(DialogInfo info)
    {
        if (info.Buttons.Count == 0)
        {
            return false;
        }
        foreach (ButtonInfo button in info.Buttons)
        {
            string normalized = NormalizeButtonText(button.Text);
            if (!String.Equals(normalized, "ok", StringComparison.Ordinal) &&
                !String.Equals(normalized, "help", StringComparison.Ordinal))
            {
                return false;
            }
        }
        return FindButton(info, "OK") != null;
    }

    private static ButtonInfo FindButton(DialogInfo info, string caption)
    {
        string normalizedCaption = NormalizeButtonText(caption);
        foreach (ButtonInfo button in info.Buttons)
        {
            if (String.Equals(NormalizeButtonText(button.Text), normalizedCaption, StringComparison.Ordinal))
            {
                return button;
            }
        }
        return null;
    }

    private static string NormalizeButtonText(string text)
    {
        return (text ?? "").Replace("&", "").Trim().ToLowerInvariant();
    }

    private static string PickMessage(DialogInfo info)
    {
        string best = "";
        foreach (string text in info.Texts)
        {
            if (text.Length > best.Length)
            {
                best = text;
            }
        }
        return best;
    }

    private static string DialogKey(DialogInfo info)
    {
        return info.Handle.ToInt64().ToString() + "|" + info.Title + "|" + info.Message + "|" + ButtonCaptions(info);
    }

    private static string ButtonCaptions(DialogInfo info)
    {
        List<string> captions = new List<string>();
        foreach (ButtonInfo button in info.Buttons)
        {
            captions.Add(button.Text);
        }
        return Join(captions, "|");
    }

    private static bool ClickButton(IntPtr buttonHandle)
    {
        IntPtr result;
        return SendMessageTimeout(
            buttonHandle,
            BM_CLICK,
            IntPtr.Zero,
            IntPtr.Zero,
            SMTO_ABORTIFHUNG,
            500,
            out result) != IntPtr.Zero;
    }

    private static string WindowClass(IntPtr hWnd)
    {
        StringBuilder builder = new StringBuilder(256);
        int length = GetClassName(hWnd, builder, builder.Capacity);
        return length > 0 ? builder.ToString(0, length) : "";
    }

    private static string WindowText(IntPtr hWnd)
    {
        int length = GetWindowTextLength(hWnd);
        StringBuilder builder = new StringBuilder(Math.Max(1, length + 1));
        int copied = GetWindowText(hWnd, builder, builder.Capacity);
        return copied > 0 ? builder.ToString(0, copied) : "";
    }

    private static void EmitDetected(DialogInfo info, DialogAction action)
    {
        StringBuilder json = BaseEvent("modal-detected");
        AddString(json, "title", info.Title);
        AddString(json, "className", info.ClassName);
        AddString(json, "message", info.Message);
        AddStringArray(json, "texts", info.Texts);
        AddButtonArray(json, info);
        AddBool(json, "safeToDismiss", action.SafeToDismiss);
        AddString(json, "classification", action.Classification);
        FinishEvent(json);
    }

    private static void EmitDismissed(DialogInfo info, string button, bool dismissed)
    {
        StringBuilder json = BaseEvent("modal-dismissed");
        AddString(json, "title", info.Title);
        AddString(json, "message", info.Message);
        AddString(json, "button", button);
        AddBool(json, "dismissed", dismissed);
        FinishEvent(json);
    }

    private static void EmitBlockedOnce(DialogInfo info, DialogAction action, string key)
    {
        bool shouldEmit;
        lock (Gate)
        {
            shouldEmit = Seen.Add("blocked|" + key);
        }
        if (!shouldEmit)
        {
            return;
        }

        StringBuilder json = BaseEvent("modal-blocked");
        AddString(json, "title", info.Title);
        AddString(json, "message", info.Message);
        AddButtonArray(json, info);
        AddString(json, "reason", action.Reason);
        FinishEvent(json);
    }

    private static StringBuilder BaseEvent(string kind)
    {
        string prefix;
        string excelId;
        string qualifiedName;
        lock (Gate)
        {
            prefix = EventPrefix;
            excelId = ExcelId;
            qualifiedName = QualifiedName;
        }

        StringBuilder json = new StringBuilder();
        json.Append(prefix);
        json.Append("{");
        AddString(json, "kind", kind);
        AddString(json, "excelId", excelId);
        AddString(json, "qualifiedName", qualifiedName);
        return json;
    }

    private static void FinishEvent(StringBuilder json)
    {
        json.Append("}");
        Console.Out.WriteLine(json.ToString());
        Console.Out.Flush();
    }

    private static void AddString(StringBuilder json, string name, string value)
    {
        if (value == null)
        {
            value = "";
        }
        AddSeparatorIfNeeded(json);
        json.Append("\"");
        json.Append(JsonEscape(name));
        json.Append("\":\"");
        json.Append(JsonEscape(value));
        json.Append("\"");
    }

    private static void AddBool(StringBuilder json, string name, bool value)
    {
        AddSeparatorIfNeeded(json);
        json.Append("\"");
        json.Append(JsonEscape(name));
        json.Append("\":");
        json.Append(value ? "true" : "false");
    }

    private static void AddStringArray(StringBuilder json, string name, List<string> values)
    {
        AddSeparatorIfNeeded(json);
        json.Append("\"");
        json.Append(JsonEscape(name));
        json.Append("\":[");
        for (int i = 0; i < values.Count; i++)
        {
            if (i > 0)
            {
                json.Append(",");
            }
            json.Append("\"");
            json.Append(JsonEscape(values[i]));
            json.Append("\"");
        }
        json.Append("]");
    }

    private static void AddButtonArray(StringBuilder json, DialogInfo info)
    {
        List<string> buttons = new List<string>();
        foreach (ButtonInfo button in info.Buttons)
        {
            buttons.Add(button.Text);
        }
        AddStringArray(json, "buttons", buttons);
    }

    private static void AddSeparatorIfNeeded(StringBuilder json)
    {
        if (json.Length == 0)
        {
            return;
        }
        char last = json[json.Length - 1];
        if (last != '{' && last != '[' && last != ',')
        {
            json.Append(",");
        }
    }

    private static string JsonEscape(string value)
    {
        StringBuilder escaped = new StringBuilder();
        foreach (char ch in value)
        {
            switch (ch)
            {
                case '\\':
                    escaped.Append("\\\\");
                    break;
                case '"':
                    escaped.Append("\\\"");
                    break;
                case '\r':
                    escaped.Append("\\r");
                    break;
                case '\n':
                    escaped.Append("\\n");
                    break;
                case '\t':
                    escaped.Append("\\t");
                    break;
                default:
                    escaped.Append(ch < ' ' ? ' ' : ch);
                    break;
            }
        }
        return escaped.ToString();
    }

    private static string Join(List<string> values, string separator)
    {
        StringBuilder joined = new StringBuilder();
        for (int i = 0; i < values.Count; i++)
        {
            if (i > 0)
            {
                joined.Append(separator);
            }
            joined.Append(values[i]);
        }
        return joined.ToString();
    }
}
`;
}

export function buildOwnedReadOnlyExcelTestHostScript(
    filePath: string,
    tests: readonly VbaTestHostPlanItem[],
    options: OwnedReadOnlyExcelTestHostScriptOptions = {},
): string {
    const testsJson = JSON.stringify(tests);
    const modalWatcherSource = productionModalWatcherCSharp();
    return [
        '$ErrorActionPreference = "Stop"',
        '$ProgressPreference = "SilentlyContinue"',
        `$targetPath = ${psSingleQuoted(filePath)}`,
        `$testsJson = ${psSingleQuoted(testsJson)}`,
        `$failFast = ${options.failFast ? '$true' : '$false'}`,
        `$eventPrefix = ${psSingleQuoted(XLIDE_TEST_HOST_EVENT_PREFIX)}`,
        '$excelId = "xlide-" + [Guid]::NewGuid().ToString("N")',
        '$excel = $null',
        '$workbook = $null',
        '$modalWatcherAvailable = $false',
        "$pidHelperSource = 'using System; using System.Runtime.InteropServices; public static class XlideWin32 { [DllImport(\"user32.dll\")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId); }'",
        `$modalWatcherSource = ${psSingleQuoted(modalWatcherSource)}`,
        'try { Add-Type -TypeDefinition $pidHelperSource -ErrorAction SilentlyContinue } catch { }',
        'try { Add-Type -TypeDefinition $modalWatcherSource -ErrorAction Stop; $modalWatcherAvailable = $true } catch { $modalWatcherAvailable = $false }',
        'function Emit-XlideTestHostEvent([string]$kind, [hashtable]$payload) {',
        '  $payload["kind"] = $kind',
        '  [Console]::Out.WriteLine($eventPrefix + ($payload | ConvertTo-Json -Compress -Depth 8))',
        '  [Console]::Out.Flush()',
        '}',
        'try {',
        '  $tests = ConvertFrom-Json -InputObject $testsJson',
        '  $excel = New-Object -ComObject Excel.Application',
        '  $excel.Visible = $false',
        '  $excel.DisplayAlerts = $false',
        '  try { $excel.AskToUpdateLinks = $false } catch { }',
        '  try { $excel.EnableEvents = $false } catch { }',
        '  try { $excel.ScreenUpdating = $false } catch { }',
        '  $excelPid = $null',
        '  try {',
        '    $processId = [uint32]0',
        '    [void][XlideWin32]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd, [ref]$processId)',
        '    if ($processId -gt 0) { $excelPid = [int]$processId }',
        '  } catch { }',
        '  Emit-XlideTestHostEvent "excel-created" @{ excelId = $excelId; owned = $true; pid = $excelPid; visible = $false }',
        '  try {',
        '    $workbook = $excel.Workbooks.Open($targetPath, 0, $true, [Type]::Missing, [Type]::Missing, [Type]::Missing, $true)',
        '  } catch {',
        '    throw ("OPEN_FAILED|XLIDE could not open the workbook read-only for tests: " + $_.Exception.Message)',
        '  }',
        '  Emit-XlideTestHostEvent "workbook-opened" @{ excelId = $excelId; filePath = $targetPath; readOnly = $true; updateLinks = 0; displayAlerts = $false; ignoreReadOnlyRecommended = $true }',
        '  foreach ($test in @($tests)) {',
        '    $macroName = [string]$test.qualifiedName',
        '    $timeoutMs = [int]$test.timeoutMs',
        '    $expectedFailure = [bool]$test.expectedFailure',
        '    Emit-XlideTestHostEvent "macro-started" @{ excelId = $excelId; qualifiedName = $macroName; timeoutMs = $timeoutMs }',
        '    $sw = [Diagnostics.Stopwatch]::StartNew()',
        '    try {',
        '      $macroRef = "\'" + $workbook.Name + "\'!" + $macroName',
        '      if ($modalWatcherAvailable -and $excelPid) { [XlideTestModalWatcher]::Start([uint32]$excelPid, $eventPrefix, $excelId, $macroName) }',
        '      $excel.Run($macroRef)',
        '      $sw.Stop()',
        '      Emit-XlideTestHostEvent "macro-finished" @{ excelId = $excelId; qualifiedName = $macroName; outcome = "passed"; durationMs = [int]$sw.ElapsedMilliseconds }',
        '      if ($failFast -and $expectedFailure) { break }',
        '    } catch {',
        '      $sw.Stop()',
        '      $message = $_.Exception.Message',
        '      Emit-XlideTestHostEvent "macro-finished" @{ excelId = $excelId; qualifiedName = $macroName; outcome = "failed"; durationMs = [int]$sw.ElapsedMilliseconds; message = $message }',
        '      if ($failFast -and -not $expectedFailure) { break }',
        '    } finally {',
        '      if ($modalWatcherAvailable) { try { [XlideTestModalWatcher]::Stop() } catch { } }',
        '    }',
        '  }',
        '} catch {',
        '  [Console]::Error.WriteLine("XLIDE_TEST_HOST_ERROR|" + $_.Exception.Message)',
        '  exit 1',
        '} finally {',
        '  if ($modalWatcherAvailable) { try { [XlideTestModalWatcher]::Stop() } catch { } }',
        '  if ($workbook) {',
        '    try { $workbook.Close($false); Emit-XlideTestHostEvent "workbook-closed" @{ excelId = $excelId; filePath = $targetPath; saveChanges = $false } } catch { }',
        '    try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) } catch { }',
        '  }',
        '  if ($excel) {',
        '    try { $excel.Quit(); Emit-XlideTestHostEvent "excel-quit" @{ excelId = $excelId } } catch { }',
        '    try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel) } catch { }',
        '  }',
        '  try { [GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect(); [GC]::WaitForPendingFinalizers() } catch { }',
        '}',
        'exit 0',
    ].join('; ');
}

export function parseVbaTestHostEventLine(line: string): VbaTestHostOracleEvent | undefined {
    if (!line.startsWith(XLIDE_TEST_HOST_EVENT_PREFIX)) {
        return undefined;
    }
    const json = line.slice(XLIDE_TEST_HOST_EVENT_PREFIX.length);
    const parsed = JSON.parse(json) as VbaTestHostOracleEvent;
    return parsed;
}

function psSingleQuoted(value: string): string {
    return `'${value.replace(/'/g, "''")}'`;
}
