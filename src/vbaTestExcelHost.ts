import type { VbaTestCase } from './vbaTestRunner';
import type { VbaTestHostOracleEvent } from './vbaTestHostOracle';
import { readExtensionTextAsset } from './extensionAssets';
import { XLIDE_TEST_RUNNER_MODULE_NAME } from './vbaTestRunnerModuleCodegen';
import { psSingleQuoted } from './util/powershell';

export const XLIDE_TEST_HOST_EVENT_PREFIX = 'XLIDE_TEST_HOST_EVENT|';
export const DEFAULT_VBA_TEST_TIMEOUT_MS = 30000;

export interface VbaTestHostPlanItem {
    qualifiedName: string;
    timeoutMs: number;
    expectedFailure: boolean;
}

export interface OwnedReadOnlyExcelTestHostScriptOptions {
    failFast?: boolean;
    runnerModuleName?: string;
}

export function vbaTestHostPlanItems(tests: readonly VbaTestCase[]): VbaTestHostPlanItem[] {
    return tests.map((test) => ({
        qualifiedName: test.qualifiedName,
        timeoutMs: test.metadata.timeoutMs ?? DEFAULT_VBA_TEST_TIMEOUT_MS,
        expectedFailure: Boolean(test.metadata.xfailReason),
    }));
}

// The C# modal watcher is a pure static asset: it ships as
// assets/testhost/XlideTestModalWatcher.cs and is read from the installed
// extension layout at runtime (see extensionAssets.ts).
function productionModalWatcherCSharp(): string {
    return readExtensionTextAsset('assets/testhost/XlideTestModalWatcher.cs');
}

export function buildOwnedReadOnlyExcelTestHostScript(
    filePath: string,
    tests: readonly VbaTestHostPlanItem[],
    options: OwnedReadOnlyExcelTestHostScriptOptions = {},
): string {
    const testsJson = JSON.stringify(tests);
    const runnerModuleName = options.runnerModuleName ?? XLIDE_TEST_RUNNER_MODULE_NAME;
    const modalWatcherSource = productionModalWatcherCSharp();
    return [
        '$ErrorActionPreference = "Stop"',
        '$ProgressPreference = "SilentlyContinue"',
        `$targetPath = ${psSingleQuoted(filePath)}`,
        `$testsJson = ${psSingleQuoted(testsJson)}`,
        `$runnerModuleName = ${psSingleQuoted(runnerModuleName)}`,
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
        'function Emit-XlideHostPhase([string]$phase, [string]$outcome, [int]$durationMs, [string]$message = $null) {',
        '  $payload = @{ excelId = $excelId; phase = $phase; outcome = $outcome; durationMs = $durationMs }',
        '  if ($message) { $payload["message"] = $message }',
        '  Emit-XlideTestHostEvent "host-phase" $payload',
        '}',
        'function Format-XlideHResult([object]$hresult) {',
        '  try {',
        '    $signed = [int64]$hresult',
        '    $unsigned = if ($signed -lt 0) { [uint32]($signed + 4294967296) } else { [uint32]$signed }',
        '    return ("0x{0:X8}" -f $unsigned)',
        '  } catch {',
        '    return [string]$hresult',
        '  }',
        '}',
        'function Convert-XlideVbaRunResult([string]$json) {',
        '  try {',
        '    return ConvertFrom-Json -InputObject $json -ErrorAction Stop',
        '  } catch {',
        '    return [pscustomobject]@{ outcome = "failed"; number = 0; source = "XLIDE"; message = "XLIDE test support returned an unreadable result." }',
        '  }',
        '}',
        'function Format-XlideVbaRunResult([object]$result) {',
        '  $message = [string]$result.message',
        '  $number = [string]$result.number',
        '  $source = [string]$result.source',
        '  if ([string]::IsNullOrWhiteSpace($message)) { $message = "VBA test failed." }',
        '  if ($source -eq "XLIDE.Assert") { return $message }',
        '  if (-not [string]::IsNullOrWhiteSpace($number) -and $number -ne "0") {',
        '    $prefix = "VBA error " + $number',
        '    if (-not [string]::IsNullOrWhiteSpace($source)) { $prefix = $prefix + " from " + $source }',
        '    return $prefix + ": " + $message',
        '  }',
        '  return $message',
        '}',
        'function Convert-XlideNullableInt([object]$value) {',
        '  try {',
        '    if ($null -eq $value) { return $null }',
        '    $text = [string]$value',
        '    if ([string]::IsNullOrWhiteSpace($text)) { return $null }',
        '    return [int]$text',
        '  } catch {',
        '    return $null',
        '  }',
        '}',
        'function Convert-XlideOutputLines([object]$value) {',
        '  $lines = New-Object System.Collections.Generic.List[string]',
        '  if ($null -eq $value) { return $lines.ToArray() }',
        '  foreach ($item in @($value)) {',
        '    if ($null -ne $item) { [void]$lines.Add([string]$item) }',
        '  }',
        '  return $lines.ToArray()',
        '}',
        // Emits raw HRESULT and exception text only; the friendly wording for
        // known HRESULTs is owned by the TypeScript side (vbaTestFailureMessages.ts).
        'function Format-XlideRunException([object]$errorRecord) {',
        '  $lines = New-Object System.Collections.Generic.List[string]',
        '  $exception = $errorRecord.Exception',
        '  if ($exception) {',
        '    $message = [string]$exception.Message',
        '    if ($exception.HResult -ne 0) {',
        '      $hex = Format-XlideHResult $exception.HResult',
        '      [void]$lines.Add("HRESULT: " + $hex)',
        '    }',
        '    if ($message) { [void]$lines.Add($message) }',
        '    $inner = $exception.InnerException',
        '    while ($inner) {',
        '      if ($inner.Message) { [void]$lines.Add("Inner: " + [string]$inner.Message) }',
        '      if ($inner.HResult -ne 0) { [void]$lines.Add("Inner HRESULT: " + (Format-XlideHResult $inner.HResult)) }',
        '      $inner = $inner.InnerException',
        '    }',
        '  }',
        '  if ($lines.Count -eq 0) { return [string]$errorRecord }',
        '  return ($lines | Select-Object -Unique) -join [Environment]::NewLine',
        '}',
        'try {',
        '  $tests = ConvertFrom-Json -InputObject $testsJson',
        '  $phaseSw = [Diagnostics.Stopwatch]::StartNew()',
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
        '  $phaseSw.Stop()',
        '  Emit-XlideHostPhase "excel-create" "passed" ([int]$phaseSw.ElapsedMilliseconds)',
        '  Emit-XlideTestHostEvent "excel-created" @{ excelId = $excelId; owned = $true; pid = $excelPid; visible = $false }',
        '  $phaseSw = [Diagnostics.Stopwatch]::StartNew()',
        '  try {',
        '    $workbook = $excel.Workbooks.Open($targetPath, 0, $true, [Type]::Missing, [Type]::Missing, [Type]::Missing, $true)',
        '  } catch {',
        '    $phaseSw.Stop()',
        '    $openMessage = "OPEN_FAILED|XLIDE could not open the workbook read-only for tests: " + $_.Exception.Message',
        '    Emit-XlideHostPhase "workbook-open" "failed" ([int]$phaseSw.ElapsedMilliseconds) $openMessage',
        '    throw $openMessage',
        '  }',
        '  $phaseSw.Stop()',
        '  Emit-XlideHostPhase "workbook-open" "passed" ([int]$phaseSw.ElapsedMilliseconds)',
        '  Emit-XlideTestHostEvent "workbook-opened" @{ excelId = $excelId; filePath = $targetPath; readOnly = $true; updateLinks = 0; displayAlerts = $false; ignoreReadOnlyRecommended = $true }',
        '  foreach ($test in @($tests)) {',
        '    $macroName = [string]$test.qualifiedName',
        '    $timeoutMs = [int]$test.timeoutMs',
        '    $expectedFailure = [bool]$test.expectedFailure',
        '    Emit-XlideTestHostEvent "macro-started" @{ excelId = $excelId; qualifiedName = $macroName; timeoutMs = $timeoutMs }',
        '    $sw = [Diagnostics.Stopwatch]::StartNew()',
        '    try {',
        '      $testRunnerRef = "\'" + $workbook.Name + "\'!" + $runnerModuleName + ".RunTest"',
        '      if ($modalWatcherAvailable -and $excelPid) { [XlideTestModalWatcher]::Start([uint32]$excelPid, $eventPrefix, $excelId, $macroName) }',
        '      $vbaRunResult = Convert-XlideVbaRunResult ([string]$excel.Run($testRunnerRef, $macroName))',
        '      $testOutput = Convert-XlideOutputLines $vbaRunResult.output',
        '      $sw.Stop()',
        '      if ([string]$vbaRunResult.outcome -eq "passed") {',
        '        $payload = @{ excelId = $excelId; qualifiedName = $macroName; outcome = "passed"; durationMs = [int]$sw.ElapsedMilliseconds }',
        '        if ($testOutput.Count -gt 0) { $payload["output"] = @($testOutput) }',
        '        Emit-XlideTestHostEvent "macro-finished" $payload',
        '        if ($failFast -and $expectedFailure) { break }',
        '      } else {',
        '        $message = "RUN_FAILED|" + (Format-XlideVbaRunResult $vbaRunResult)',
        '        $payload = @{ excelId = $excelId; qualifiedName = $macroName; outcome = "failed"; durationMs = [int]$sw.ElapsedMilliseconds; message = $message }',
        '        $errorNumber = Convert-XlideNullableInt $vbaRunResult.number',
        '        if ($null -ne $errorNumber) { $payload["errorNumber"] = $errorNumber }',
        '        $errorSource = [string]$vbaRunResult.source',
        '        if (-not [string]::IsNullOrWhiteSpace($errorSource)) { $payload["errorSource"] = $errorSource }',
        '        if ($testOutput.Count -gt 0) { $payload["output"] = @($testOutput) }',
        '        Emit-XlideTestHostEvent "macro-finished" $payload',
        '        if ($failFast -and -not $expectedFailure) { break }',
        '      }',
        '    } catch {',
        '      $sw.Stop()',
        '      $message = "RUN_FAILED|" + (Format-XlideRunException $_)',
        '      Emit-XlideTestHostEvent "macro-finished" @{ excelId = $excelId; qualifiedName = $macroName; outcome = "runner-error"; durationMs = [int]$sw.ElapsedMilliseconds; message = $message }',
        '      break',
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
        '    $phaseSw = [Diagnostics.Stopwatch]::StartNew()',
        '    try {',
        '      $workbook.Close($false)',
        '      $phaseSw.Stop()',
        '      Emit-XlideTestHostEvent "workbook-closed" @{ excelId = $excelId; filePath = $targetPath; saveChanges = $false; durationMs = [int]$phaseSw.ElapsedMilliseconds }',
        '      Emit-XlideHostPhase "workbook-close" "passed" ([int]$phaseSw.ElapsedMilliseconds)',
        '    } catch {',
        '      $phaseSw.Stop()',
        '      Emit-XlideHostPhase "workbook-close" "failed" ([int]$phaseSw.ElapsedMilliseconds) $_.Exception.Message',
        '    }',
        '  }',
        '  if ($excel) {',
        '    $phaseSw = [Diagnostics.Stopwatch]::StartNew()',
        '    try {',
        '      $excel.Quit()',
        '      $phaseSw.Stop()',
        '      Emit-XlideTestHostEvent "excel-quit" @{ excelId = $excelId; durationMs = [int]$phaseSw.ElapsedMilliseconds }',
        '      Emit-XlideHostPhase "excel-quit" "passed" ([int]$phaseSw.ElapsedMilliseconds)',
        '    } catch {',
        '      $phaseSw.Stop()',
        '      Emit-XlideHostPhase "excel-quit" "failed" ([int]$phaseSw.ElapsedMilliseconds) $_.Exception.Message',
        '    }',
        '  }',
        '  $releaseSw = [Diagnostics.Stopwatch]::StartNew()',
        '  try { if ($workbook) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) } } catch { }',
        '  try { if ($excel) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel) } } catch { }',
        '  try { [GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect(); [GC]::WaitForPendingFinalizers() } catch { }',
        '  $releaseSw.Stop()',
        '  Emit-XlideHostPhase "com-release" "passed" ([int]$releaseSw.ElapsedMilliseconds)',
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
