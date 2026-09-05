# XLIDE owned read-only Office test host (static body).
# buildOwnedReadOnlyExcelTestHostScript (src/vbaTestExcelHost.ts) prepends the
# dynamic preamble before this body: $ErrorActionPreference,
# $ProgressPreference, $targetPath, $testsJson, $runnerModuleName, $failFast,
# $eventPrefix, $hostKind/$hostProgId/$hostProcessName/$hostNoun (which Office
# application hosts the run: excel, word, or powerpoint), and
# $modalWatcherSource (assets/testhost/XlideTestModalWatcher.cs).
# Event names keep their excel-flavored identifiers for every host: they are
# wire-protocol ids consumed by vbaTestHostOracle, not display text.
# Emit UTF-8 so non-ASCII host/COM error text is decoded correctly by the Node side.
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$excelId = "xlide-" + [Guid]::NewGuid().ToString("N")
$excel = $null
$workbook = $null
$modalWatcherAvailable = $false
$xlideJob = $null
# GetWindowThreadProcessId resolves the owned Excel's PID. The job object ties
# that Excel's lifetime to this PowerShell process: JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE
# means the kernel kills Excel when the last handle closes, so a crashed or
# force-killed host can never orphan a hidden EXCEL.EXE holding the workbook open.
$pidHelperSource = @'
using System;
using System.Runtime.InteropServices;
public static class XlideWin32 {
  [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
  [DllImport("kernel32.dll", CharSet=CharSet.Unicode)] static extern IntPtr CreateJobObjectW(IntPtr a, string name);
  [DllImport("kernel32.dll")] static extern bool SetInformationJobObject(IntPtr job, int infoClass, IntPtr info, uint len);
  [DllImport("kernel32.dll")] static extern bool AssignProcessToJobObject(IntPtr job, IntPtr process);
  [DllImport("kernel32.dll")] static extern IntPtr OpenProcess(uint access, bool inherit, uint pid);
  [DllImport("kernel32.dll")] static extern bool CloseHandle(IntPtr h);
  const int JobObjectExtendedLimitInformation = 9;
  const int LimitKillOnJobClose = 0x2000;
  const uint ProcessSetQuota = 0x0100;
  const uint ProcessTerminate = 0x0001;
  // JOBOBJECT_EXTENDED_LIMIT_INFORMATION: LimitFlags sits at offset 16 of the
  // basic limits, which lead the struct. 144/112 bytes on x64/x86.
  public static IntPtr CreateKillOnCloseJob() {
    IntPtr job = CreateJobObjectW(IntPtr.Zero, null);
    if (job == IntPtr.Zero) { return IntPtr.Zero; }
    int size = IntPtr.Size == 8 ? 144 : 112;
    IntPtr info = Marshal.AllocHGlobal(size);
    try {
      for (int i = 0; i < size; i++) { Marshal.WriteByte(info, i, 0); }
      Marshal.WriteInt32(info, 16, LimitKillOnJobClose);
      if (!SetInformationJobObject(job, JobObjectExtendedLimitInformation, info, (uint)size)) {
        CloseHandle(job);
        return IntPtr.Zero;
      }
    } finally { Marshal.FreeHGlobal(info); }
    return job;
  }
  public static bool AssignToJob(IntPtr job, uint pid) {
    if (job == IntPtr.Zero || pid == 0) { return false; }
    IntPtr process = OpenProcess(ProcessSetQuota | ProcessTerminate, false, pid);
    if (process == IntPtr.Zero) { return false; }
    try { return AssignProcessToJobObject(job, process); }
    finally { CloseHandle(process); }
  }
}
'@
try { Add-Type -TypeDefinition $pidHelperSource -ErrorAction SilentlyContinue } catch { }
try { Add-Type -TypeDefinition $modalWatcherSource -ErrorAction Stop; $modalWatcherAvailable = $true } catch { $modalWatcherAvailable = $false }
function Emit-XlideTestHostEvent([string]$kind, [hashtable]$payload) {
  $payload["kind"] = $kind
  [Console]::Out.WriteLine($eventPrefix + ($payload | ConvertTo-Json -Compress -Depth 8))
  [Console]::Out.Flush()
}
function Emit-XlideHostPhase([string]$phase, [string]$outcome, [int]$durationMs, [string]$message = $null) {
  $payload = @{ excelId = $excelId; phase = $phase; outcome = $outcome; durationMs = $durationMs }
  if ($message) { $payload["message"] = $message }
  Emit-XlideTestHostEvent "host-phase" $payload
}
function Format-XlideHResult([object]$hresult) {
  try {
    $signed = [int64]$hresult
    $unsigned = if ($signed -lt 0) { [uint32]($signed + 4294967296) } else { [uint32]$signed }
    return ("0x{0:X8}" -f $unsigned)
  } catch {
    return [string]$hresult
  }
}
function Convert-XlideVbaRunResult([string]$json) {
  try {
    return ConvertFrom-Json -InputObject $json -ErrorAction Stop
  } catch {
    return [pscustomobject]@{ outcome = "failed"; number = 0; source = "XLIDE"; message = "XLIDE test support returned an unreadable result." }
  }
}
function Format-XlideVbaRunResult([object]$result) {
  $message = [string]$result.message
  $number = [string]$result.number
  $source = [string]$result.source
  if ([string]::IsNullOrWhiteSpace($message)) { $message = "VBA test failed." }
  if ($source -eq "XLIDE.Assert") { return $message }
  if (-not [string]::IsNullOrWhiteSpace($number) -and $number -ne "0") {
    $prefix = "VBA error " + $number
    if (-not [string]::IsNullOrWhiteSpace($source)) { $prefix = $prefix + " from " + $source }
    return $prefix + ": " + $message
  }
  return $message
}
function Convert-XlideNullableInt([object]$value) {
  try {
    if ($null -eq $value) { return $null }
    $text = [string]$value
    if ([string]::IsNullOrWhiteSpace($text)) { return $null }
    return [int]$text
  } catch {
    return $null
  }
}
function Set-XlideHostAlertsOff([object]$app) {
  if ($null -eq $app) { return }
  try {
    if ($hostKind -eq "word") { $app.DisplayAlerts = 0 }
    elseif ($hostKind -eq "powerpoint") { $app.DisplayAlerts = 1 }
    elseif ($hostKind -eq "access") { $app.SetOption("Confirm Action Queries", $false) }
    else { $app.DisplayAlerts = $false }
  } catch { }
}
function Convert-XlideOutputLines([object]$value) {
  $lines = New-Object System.Collections.Generic.List[string]
  if ($null -eq $value) { return $lines.ToArray() }
  foreach ($item in @($value)) {
    if ($null -ne $item) { [void]$lines.Add([string]$item) }
  }
  return $lines.ToArray()
}
# Emits raw HRESULT and exception text only; the friendly wording for
# known HRESULTs is owned by the TypeScript side (vbaTestFailureMessages.ts).
function Format-XlideRunException([object]$errorRecord) {
  $lines = New-Object System.Collections.Generic.List[string]
  $exception = $errorRecord.Exception
  if ($exception) {
    $message = [string]$exception.Message
    if ($exception.HResult -ne 0) {
      $hex = Format-XlideHResult $exception.HResult
      [void]$lines.Add("HRESULT: " + $hex)
    }
    if ($message) { [void]$lines.Add($message) }
    $inner = $exception.InnerException
    while ($inner) {
      if ($inner.Message) { [void]$lines.Add("Inner: " + [string]$inner.Message) }
      if ($inner.HResult -ne 0) { [void]$lines.Add("Inner HRESULT: " + (Format-XlideHResult $inner.HResult)) }
      $inner = $inner.InnerException
    }
  }
  if ($lines.Count -eq 0) { return [string]$errorRecord }
  return ($lines | Select-Object -Unique) -join [Environment]::NewLine
}
try {
  $tests = ConvertFrom-Json -InputObject $testsJson
  $phaseSw = [Diagnostics.Stopwatch]::StartNew()
  # Snapshot the Excels already running so ownership can be PROVEN, not assumed.
  # This host quits its instance on the way out and kills it on a hang, so
  # attaching to a user's Excel would put their unsaved work in the blast radius.
  $preExistingExcelPids = @()
  try { $preExistingExcelPids = @(Get-Process -Name $hostProcessName -ErrorAction SilentlyContinue | ForEach-Object { [int]$_.Id }) } catch { }
  $excel = New-Object -ComObject $hostProgId
  # Ownership is proven BEFORE any setting is touched: a single-instance host
  # (PowerPoint) hands back the user's own running application here, and that
  # instance must be refused with its alerts and macro security untouched.
  $excelPid = $null
  try {
    # Excel and PowerPoint expose a window handle; Word does not, and falls
    # through to the process-diff below.
    $hostHwnd = [IntPtr]::Zero
    if ($hostKind -eq "excel") { $hostHwnd = [IntPtr]$excel.Hwnd }
    elseif ($hostKind -eq "powerpoint") { $hostHwnd = [IntPtr]$excel.HWND }
    elseif ($hostKind -eq "access") { $hostHwnd = [IntPtr]$excel.hWndAccessApp() }
    if ($hostHwnd -ne [IntPtr]::Zero) {
      $processId = [uint32]0
      [void][XlideWin32]::GetWindowThreadProcessId($hostHwnd, [ref]$processId)
      if ($processId -gt 0) { $excelPid = [int]$processId }
    }
  } catch { }
  if (-not $excelPid) {
    try {
      $afterPids = @(Get-Process -Name $hostProcessName -ErrorAction SilentlyContinue | ForEach-Object { [int]$_.Id })
      $newPids = @($afterPids | Where-Object { $preExistingExcelPids -notcontains $_ })
      if ($newPids.Count -eq 1) { $excelPid = [int]$newPids[0] }
    } catch { }
  }
  if ($excelPid -and ($preExistingExcelPids -contains $excelPid)) {
    # Refuse rather than proceed: this Excel is someone else's.
    $excel = $null
    $phaseSw.Stop()
    $ownershipMessage = "XLIDE refused to run tests: the new " + $hostNoun + " Application resolved to already-running process " + $excelPid + ". The test host must own its " + $hostNoun + " instance because it quits that process when the run ends. Close the running " + $hostNoun + " and try again."
    Emit-XlideHostPhase "excel-create" "failed" ([int]$phaseSw.ElapsedMilliseconds) $ownershipMessage
    throw $ownershipMessage
  }
  # Owned and proven: configure the instance for silent automation.
  # PowerPoint refuses Application.Visible = False outright; its run stays
  # windowless anyway because the presentation opens WithWindow:=msoFalse.
  if ($hostKind -ne "powerpoint") { try { $excel.Visible = $false } catch { } }
  Set-XlideHostAlertsOff $excel
  try { $excel.AskToUpdateLinks = $false } catch { }
  try { $excel.EnableEvents = $false } catch { }
  try { $excel.ScreenUpdating = $false } catch { }
  # msoAutomationSecurityLow (1): open workbooks without the macro-security
  # prompt, which is a host-owned modal that no watcher can dismiss.
  try { $excel.AutomationSecurity = 1 } catch { }
  $jobActive = $false
  if ($excelPid) {
    try {
      $xlideJob = [XlideWin32]::CreateKillOnCloseJob()
      $jobActive = [XlideWin32]::AssignToJob($xlideJob, [uint32]$excelPid)
    } catch { $jobActive = $false }
  }
  $phaseSw.Stop()
  Emit-XlideHostPhase "excel-create" "passed" ([int]$phaseSw.ElapsedMilliseconds)
  Emit-XlideTestHostEvent "excel-created" @{ excelId = $excelId; owned = $true; pid = $excelPid; visible = $false; killOnClose = $jobActive }
  # Watch for modals from here on, not just around macro execution: opening a
  # workbook and closing it can both raise dialogs, and an unwatched dialog
  # wedges the host until its timeout instead of being reported and dismissed.
  if ($modalWatcherAvailable -and $excelPid) {
    try { [XlideTestModalWatcher]::Start([uint32]$excelPid, $eventPrefix, $excelId, "workbook-open") } catch { }
  }
  $phaseSw = [Diagnostics.Stopwatch]::StartNew()
  try {
    if ($hostKind -eq "word") {
      # Documents.Open(FileName, ConfirmConversions, ReadOnly, AddToRecentFiles)
      $workbook = $excel.Documents.Open($targetPath, $false, $true, $false)
    } elseif ($hostKind -eq "access") {
      # Access holds one database at a time and hands it back as
      # CurrentProject rather than as an open document object.
      $excel.OpenCurrentDatabase($targetPath)
      $workbook = $excel.CurrentProject
    } elseif ($hostKind -eq "powerpoint") {
      # Presentations.Open(FileName, ReadOnly:=msoTrue, Untitled:=msoFalse,
      # WithWindow:=msoFalse) - windowless, which is also why the visible
      # PowerPoint application never shows a document.
      $workbook = $excel.Presentations.Open($targetPath, -1, 0, 0)
    } else {
      $workbook = $excel.Workbooks.Open($targetPath, 0, $true, [Type]::Missing, [Type]::Missing, [Type]::Missing, $true)
    }
  } catch {
    $phaseSw.Stop()
    $openMessage = "OPEN_FAILED|XLIDE could not open the file read-only for tests: " + $_.Exception.Message
    Emit-XlideHostPhase "workbook-open" "failed" ([int]$phaseSw.ElapsedMilliseconds) $openMessage
    throw $openMessage
  }
  $phaseSw.Stop()
  Emit-XlideHostPhase "workbook-open" "passed" ([int]$phaseSw.ElapsedMilliseconds)
  Emit-XlideTestHostEvent "workbook-opened" @{ excelId = $excelId; filePath = $targetPath; readOnly = $true; updateLinks = 0; displayAlerts = $false; ignoreReadOnlyRecommended = $true }
  foreach ($test in @($tests)) {
    $macroName = [string]$test.qualifiedName
    $timeoutMs = [int]$test.timeoutMs
    $expectedFailure = [bool]$test.expectedFailure
    Emit-XlideTestHostEvent "macro-started" @{ excelId = $excelId; qualifiedName = $macroName; timeoutMs = $timeoutMs }
    $sw = [Diagnostics.Stopwatch]::StartNew()
    try {
      if ($hostKind -eq "word") {
        # Word resolves Module.Proc and rejects document-qualified names.
        $testRunnerRef = $runnerModuleName + ".RunTest"
      } elseif ($hostKind -eq "access") {
        # Access takes the bare procedure name: measured on 16.0, where
        # Module1.Main is refused with "cannot find the procedure".
        $testRunnerRef = "RunTest"
      } elseif ($hostKind -eq "powerpoint") {
        # PowerPoint takes the presentation-qualified File.pptm!Module.Proc form.
        $testRunnerRef = $workbook.Name + "!" + $runnerModuleName + ".RunTest"
      } else {
        $testRunnerRef = "'" + ($workbook.Name -replace "'", "''") + "'!" + $runnerModuleName + ".RunTest"
      }
      if ($modalWatcherAvailable -and $excelPid) { [XlideTestModalWatcher]::Start([uint32]$excelPid, $eventPrefix, $excelId, $macroName) }
      if ($hostKind -eq "word" -or $hostKind -eq "access") {
        # Word and Access both declare Run's varargs ByRef, so PowerShell COM
        # interop requires a [ref] wrapper.
        $macroArg = $macroName
        $vbaRunResult = Convert-XlideVbaRunResult ([string]$excel.Run($testRunnerRef, [ref]$macroArg))
      } elseif ($hostKind -eq "powerpoint") {
        # PowerPoint's Run takes a ParamArray that rejects PowerShell's
        # ByRef-marshaled arguments ("type must not be ByRef"); reflection
        # InvokeMember marshals them ByVal.
        $vbaRunResult = Convert-XlideVbaRunResult ([string]$excel.GetType().InvokeMember("Run", [Reflection.BindingFlags]::InvokeMethod, $null, $excel, @($testRunnerRef, $macroName)))
      } else {
        $vbaRunResult = Convert-XlideVbaRunResult ([string]$excel.Run($testRunnerRef, $macroName))
      }
      $testOutput = Convert-XlideOutputLines $vbaRunResult.output
      $sw.Stop()
      if ([string]$vbaRunResult.outcome -eq "passed") {
        $payload = @{ excelId = $excelId; qualifiedName = $macroName; outcome = "passed"; durationMs = [int]$sw.ElapsedMilliseconds }
        if ($testOutput.Count -gt 0) { $payload["output"] = @($testOutput) }
        Emit-XlideTestHostEvent "macro-finished" $payload
        if ($failFast -and $expectedFailure) { break }
      } else {
        $message = "RUN_FAILED|" + (Format-XlideVbaRunResult $vbaRunResult)
        $payload = @{ excelId = $excelId; qualifiedName = $macroName; outcome = "failed"; durationMs = [int]$sw.ElapsedMilliseconds; message = $message }
        $errorNumber = Convert-XlideNullableInt $vbaRunResult.number
        if ($null -ne $errorNumber) { $payload["errorNumber"] = $errorNumber }
        $errorSource = [string]$vbaRunResult.source
        if (-not [string]::IsNullOrWhiteSpace($errorSource)) { $payload["errorSource"] = $errorSource }
        if ($testOutput.Count -gt 0) { $payload["output"] = @($testOutput) }
        Emit-XlideTestHostEvent "macro-finished" $payload
        if ($failFast -and -not $expectedFailure) { break }
      }
    } catch {
      $sw.Stop()
      $message = "RUN_FAILED|" + (Format-XlideRunException $_)
      Emit-XlideTestHostEvent "macro-finished" @{ excelId = $excelId; qualifiedName = $macroName; outcome = "runner-error"; durationMs = [int]$sw.ElapsedMilliseconds; message = $message }
      # Only abort the whole run on a per-test error when fail-fast is on. Otherwise
      # continue so each remaining test gets its own macro-finished event instead of
      # being reported as a spurious host-error for "no result emitted".
      if ($failFast) { break }
    } finally {
      # Hand the watcher back to teardown rather than switching it off: Close
      # and Quit can prompt too, and between tests a stray dialog would go
      # unseen. Re-Start rebinds attribution and clears the dedup set.
      if ($modalWatcherAvailable -and $excelPid) {
        try { [XlideTestModalWatcher]::Start([uint32]$excelPid, $eventPrefix, $excelId, "host-teardown") } catch { }
      }
    }
  }
} catch {
  [Console]::Error.WriteLine("XLIDE_TEST_HOST_ERROR|" + $_.Exception.Message)
  exit 1
} finally {
  # Test VBA can set Application.DisplayAlerts = True and leave it that way,
  # which would let Close or Quit raise a prompt. Re-assert suppression before
  # each, while the watcher is still running to catch anything that slips past.
  Set-XlideHostAlertsOff $excel
  if ($workbook) {
    $phaseSw = [Diagnostics.Stopwatch]::StartNew()
    try {
      if ($hostKind -eq "word") { $workbook.Close(0) }
      elseif ($hostKind -eq "access") { $excel.CloseCurrentDatabase() }
      elseif ($hostKind -eq "powerpoint") { try { $workbook.Saved = -1 } catch { }; $workbook.Close() }
      else { $workbook.Close($false) }
      $phaseSw.Stop()
      Emit-XlideTestHostEvent "workbook-closed" @{ excelId = $excelId; filePath = $targetPath; saveChanges = $false; durationMs = [int]$phaseSw.ElapsedMilliseconds }
      Emit-XlideHostPhase "workbook-close" "passed" ([int]$phaseSw.ElapsedMilliseconds)
    } catch {
      $phaseSw.Stop()
      Emit-XlideHostPhase "workbook-close" "failed" ([int]$phaseSw.ElapsedMilliseconds) $_.Exception.Message
    }
  }
  if ($excel) {
    Set-XlideHostAlertsOff $excel
    $phaseSw = [Diagnostics.Stopwatch]::StartNew()
    try {
      $excel.Quit()
      $phaseSw.Stop()
      Emit-XlideTestHostEvent "excel-quit" @{ excelId = $excelId; durationMs = [int]$phaseSw.ElapsedMilliseconds }
      Emit-XlideHostPhase "excel-quit" "passed" ([int]$phaseSw.ElapsedMilliseconds)
    } catch {
      $phaseSw.Stop()
      Emit-XlideHostPhase "excel-quit" "failed" ([int]$phaseSw.ElapsedMilliseconds) $_.Exception.Message
    }
  }
  if ($modalWatcherAvailable) { try { [XlideTestModalWatcher]::Stop() } catch { } }
  $releaseSw = [Diagnostics.Stopwatch]::StartNew()
  try { if ($workbook) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) } } catch { }
  try { if ($excel) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel) } } catch { }
  try { [GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect(); [GC]::WaitForPendingFinalizers() } catch { }
  $releaseSw.Stop()
  Emit-XlideHostPhase "com-release" "passed" ([int]$releaseSw.ElapsedMilliseconds)
}
exit 0
