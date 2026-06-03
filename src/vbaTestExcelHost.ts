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

export function buildOwnedReadOnlyExcelTestHostScript(
    filePath: string,
    tests: readonly VbaTestHostPlanItem[],
    options: OwnedReadOnlyExcelTestHostScriptOptions = {},
): string {
    const testsJson = JSON.stringify(tests);
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
        'function Emit-XlideTestHostEvent([string]$kind, [hashtable]$payload) {',
        '  $payload["kind"] = $kind',
        '  [Console]::Out.WriteLine($eventPrefix + ($payload | ConvertTo-Json -Compress -Depth 8))',
        '  [Console]::Out.Flush()',
        '}',
        'try {',
        '  $tests = ConvertFrom-Json -InputObject $testsJson',
        '  $excel = New-Object -ComObject Excel.Application',
        '  $excel.Visible = $true',
        '  $excel.DisplayAlerts = $false',
        '  try { $excel.AskToUpdateLinks = $false } catch { }',
        '  try { $excel.EnableEvents = $false } catch { }',
        '  $excelPid = $null',
        '  try {',
        '    $excelPid = @(Get-Process EXCEL | Where-Object { $_.MainWindowHandle -eq [IntPtr]$excel.Hwnd } | Select-Object -First 1 -ExpandProperty Id)[0]',
        '  } catch { }',
        '  Emit-XlideTestHostEvent "excel-created" @{ excelId = $excelId; owned = $true; pid = $excelPid }',
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
        '      $excel.Run($macroRef)',
        '      $sw.Stop()',
        '      Emit-XlideTestHostEvent "macro-finished" @{ excelId = $excelId; qualifiedName = $macroName; outcome = "passed"; durationMs = [int]$sw.ElapsedMilliseconds }',
        '      if ($failFast -and $expectedFailure) { break }',
        '    } catch {',
        '      $sw.Stop()',
        '      $message = $_.Exception.Message',
        '      Emit-XlideTestHostEvent "macro-finished" @{ excelId = $excelId; qualifiedName = $macroName; outcome = "failed"; durationMs = [int]$sw.ElapsedMilliseconds; message = $message }',
        '      if ($failFast -and -not $expectedFailure) { break }',
        '    }',
        '  }',
        '} catch {',
        '  [Console]::Error.WriteLine("XLIDE_TEST_HOST_ERROR|" + $_.Exception.Message)',
        '  exit 1',
        '} finally {',
        '  if ($workbook) {',
        '    try { $workbook.Close($false); Emit-XlideTestHostEvent "workbook-closed" @{ excelId = $excelId; filePath = $targetPath; saveChanges = $false } } catch { }',
        '  }',
        '  if ($excel) {',
        '    try { $excel.Quit(); Emit-XlideTestHostEvent "excel-quit" @{ excelId = $excelId } } catch { }',
        '    try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel) } catch { }',
        '  }',
        '}',
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
