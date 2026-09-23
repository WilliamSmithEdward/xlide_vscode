import * as path from 'path';
import { psSingleQuoted, runPowerShell } from './util/powershell';
import { OFFICE_HOST_APPS, officeHostForPath, type OfficeHostApp } from './officeHostApps';
import { LOCK_TEST_FUNCTION, unsavedWorkFunction } from './officeFileState';
import { placeAssignmentLine, restorePlaceLines, type OfficeViewPlace } from './officeViewPlace';

export type HostMacroFailureCode = 'REOPEN_BLOCKED' | 'REOPEN_FAILED' | 'RUN_FAILED' | 'UNKNOWN';

const OPEN_SENTINEL = 'XLIDE_OPEN|';

/** Reports what an open script ended up with; `$openState` starts as "opened". */
const OPEN_SENTINEL_LINE = `[Console]::Out.WriteLine("${OPEN_SENTINEL}" + $openState)`;

/**
 * "Open Read Only" on a file that is already open.
 *
 * A copy already read-only is simply reused. A copy open for EDITING is
 * closed so the lines after this reopen it read-only - it holds the lock that
 * makes every XLIDE save fail, which is the reason to ask for read-only at
 * all - unless it holds unsaved work, which is left exactly as it is and
 * reported. Before this, an open copy was reused whatever its mode, so asking
 * for read-only on a workbook already open for editing brought the editing
 * copy forward and changed nothing, and XLIDE's saves went on failing.
 */
function readOnlyReuseLines(copy: string, isReadOnly: string, close: string): string[] {
    return [
        `if (${copy} -and -not (${isReadOnly})) {`,
        `  if (Test-XlideUnsavedWork ${copy}) {`,
        '    $openState = "keptUnsaved"',
        '  }',
        '  else {',
        `    try { ${close}; ${copy} = $null } catch { $openState = "keptEditing" }`,
        '  }',
        '}',
    ];
}

/**
 * Excel after a read-only open: a read-only workbook does not lock the file,
 * so a lock that remains belongs to a copy open for editing somewhere this
 * script cannot reach - another Excel instance, or another program. The
 * window says Read-Only either way, which is exactly why it is reported.
 * Excel can hold the file a moment longer while it finishes opening.
 */
const EXCEL_LOCKED_ELSEWHERE_LINES: readonly string[] = [
    'if ($openState -eq "opened") {',
    '  $locked = Test-XlideLocked',
    '  for ($__i = 0; $locked -and $__i -lt 5; $__i++) { Start-Sleep -Milliseconds 200; $locked = Test-XlideLocked }',
    '  if ($locked) { $openState = "lockedElsewhere" }',
    '}',
];

/** Typed macro-run failure decoded from the script's CODE|message sentinels. */
export class HostMacroError extends Error {
    constructor(message: string, readonly code: HostMacroFailureCode) {
        super(message);
        this.name = 'HostMacroError';
    }
}

const MACRO_ERROR_SENTINEL = 'XLIDE_MACRO_ERROR|';

export type ExcelLaunchMode =
    | { kind: 'open'; readOnly: boolean; background?: boolean; place?: OfficeViewPlace }
    | { kind: 'macroReadOnly'; macroName: string };

export interface ExcelLaunchScriptOptions {
    filePath: string;
    attachToRunning: boolean;
    mode: ExcelLaunchMode;
}

// Attach to a running Excel instance (optional), or start one, then look for
// an already-open copy of the target workbook.
function attachLines(filePath: string, attachToRunning: boolean): string[] {
    return [
        `$targetPath = ${psSingleQuoted(filePath)}`,
        // win32 semantics regardless of host platform: this script always
        // drives Windows Excel against a Windows path.
        `$targetName = ${psSingleQuoted(path.win32.basename(filePath))}`,
        '$excel = $null',
        '$workbook = $null',
        `$attachToRunning = ${attachToRunning ? '$true' : '$false'}`,
        'if ($attachToRunning) {',
        '  try { $excel = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application") } catch { }',
        '}',
        'if (-not $excel) {',
        '  $excel = New-Object -ComObject Excel.Application',
        '}',
        '$excel.Visible = $true',
        'foreach ($wb in @($excel.Workbooks)) {',
        '  if (($wb.FullName -ieq $targetPath) -or ($wb.Name -ieq $targetName)) { $workbook = $wb; break }',
        '}',
    ];
}

// PowerShell helper that retries a COM call while the application is busy. A
// busy state (a modal dialog such as a MsgBox left open from a previous run, or
// a transient busy moment) reports as RPC_E_CALL_REJECTED (HResult -2147418111 /
// 0x80010001) or RPC_E_SERVERCALL_RETRYLATER (-2147417846 / 0x8001010A). ~3s of
// retries rides out a transient busy and gives the user a moment to dismiss a dialog.
const COM_RETRY_HELPER =
    'function Invoke-XlideCom($Action) { for ($__i = 0; $__i -le 12; $__i++) { try { return (& $Action) } catch { if (($_.Exception.HResult -eq -2147418111 -or $_.Exception.HResult -eq -2147417846 -or $_.Exception.InnerException.HResult -eq -2147418111 -or $_.Exception.InnerException.HResult -eq -2147417846 -or $_.Exception.Message -match "rejected by callee|RETRYLATER|0x80010001|0x8001010A") -and $__i -lt 12) { Start-Sleep -Milliseconds 250; continue } else { throw } } } }';

/**
 * Activate the file's window and bring the application to the foreground. Both
 * are best-effort: a busy application must not fail the launch or the run.
 * `windowHandle` is the expression each application exposes for its main
 * window, measured live on Office 16.0.
 */
function foregroundLines(windowHandle: string, activate?: string): string[] {
    return [
        ...(activate ? [`try { ${activate} } catch { }`] : []),
        'try { Add-Type -MemberDefinition \'[DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd); [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);\' -Name XlideWin32 -Namespace XlideHelper } catch { }',
        `try { [XlideHelper.XlideWin32]::ShowWindow([IntPtr]${windowHandle}, 9); [XlideHelper.XlideWin32]::SetForegroundWindow([IntPtr]${windowHandle}) } catch { }`,
    ];
}

const EXCEL_FOREGROUND_LINES: readonly string[] = foregroundLines('$excel.Hwnd', '$workbook.Activate()');

/**
 * After an open that puts back a copy XLIDE closed: where the user was in it
 * (officeViewPlace.ts). `$app` and `$file` are the application and the copy.
 */
function placeRestoreLines(host: OfficeHostApp, place: OfficeViewPlace | undefined): string[] {
    return place ? [placeAssignmentLine(place), ...restorePlaceLines(host)] : [];
}

/**
 * Single parameterized attach/open/foreground script for the Windows Excel
 * launch paths: plain open (reuses an already-open workbook) and read-only
 * macro runs (reopen read-only, foreground, run the macro, report failures
 * through the XLIDE_MACRO_ERROR stderr sentinel).
 */
export function buildExcelLaunchScript(options: ExcelLaunchScriptOptions): string {
    const { filePath, attachToRunning, mode } = options;
    if (mode.kind === 'open') {
        return [
            '$ErrorActionPreference = "Stop"',
            COM_RETRY_HELPER,
            unsavedWorkFunction('excel'),
            LOCK_TEST_FUNCTION,
            '$openState = "opened"',
            ...attachLines(filePath, attachToRunning),
            ...(mode.readOnly
                ? readOnlyReuseLines('$workbook', '[bool]$workbook.ReadOnly', 'Invoke-XlideCom { $workbook.Close($false) }')
                : []),
            'if (-not $workbook) {',
            `  $workbook = Invoke-XlideCom { $excel.Workbooks.Open($targetPath, 0, ${mode.readOnly ? '$true' : '$false'}) }`,
            '}',
            ...(mode.place ? ['$app = $excel', '$file = $workbook', ...placeRestoreLines('excel', mode.place)] : []),
            ...(mode.background ? [] : EXCEL_FOREGROUND_LINES),
            ...(mode.readOnly ? EXCEL_LOCKED_ELSEWHERE_LINES : []),
            OPEN_SENTINEL_LINE,
        ].join('\n');
    }
    return [
        '$ErrorActionPreference = "Stop"',
        COM_RETRY_HELPER,
        unsavedWorkFunction('excel'),
        'try {',
        `  $macroName = ${psSingleQuoted(mode.macroName)}`,
        ...attachLines(filePath, attachToRunning),
        'if ($workbook) {',
        '  if (-not $workbook.ReadOnly) {',
        '    throw "REOPEN_BLOCKED|Workbook is already open for editing in Excel. Close it in Excel, then press F5 again so XLIDE can reopen the saved workbook before running the macro."',
        '  }',
        // A read-only copy takes edits too - only Save As can keep them - and
        // the reopen below closes it without saving.
        '  if (Test-XlideUnsavedWork $workbook) {',
        '    throw "REOPEN_BLOCKED|The workbook is open read-only in Excel with changes that were never saved, and reopening it to run the macro would lose them. Save them under another name or close it without saving, then press F5 again."',
        '  }',
        '  try {',
        '    Invoke-XlideCom { $workbook.Close($false) }',
        '    $workbook = $null',
        '  } catch {',
        '    throw ("REOPEN_FAILED|XLIDE could not close the existing read-only workbook before running the macro: " + $_.Exception.Message)',
        '  }',
        '}',
        'try {',
        '  $workbook = Invoke-XlideCom { $excel.Workbooks.Open($targetPath, 0, $true) }',
        '} catch {',
        '  throw ("REOPEN_FAILED|XLIDE could not reopen the workbook. If it is open outside XLIDE, close it in Excel and try again: " + $_.Exception.Message)',
        '}',
        ...EXCEL_FOREGROUND_LINES,
        '$macroRef = "\'" + ($workbook.Name -replace "\'", "\'\'") + "\'!" + $macroName',
        'try {',
        '  Invoke-XlideCom { $excel.Run($macroRef) }',
        '} catch {',
        '  throw ("RUN_FAILED|XLIDE could not run the macro: " + $_.Exception.Message)',
        '}',
        '[Console]::Out.WriteLine("XLIDE_MACRO_OK")',
        '} catch {',
        '  [Console]::Error.WriteLine("XLIDE_MACRO_ERROR|" + ($_.Exception.Message -replace "[\\r\\n]+", " "))',
        '  exit 1',
        '}',
    ].join('\n');
}

/**
 * F5 macro-run scripts for the other hosts, mirroring the Excel macroReadOnly
 * flow (find an open copy, refuse edit-mode, close a stale read-only copy,
 * reopen read-only, run, report through the same XLIDE_MACRO_ERROR
 * sentinels). Every host-specific line is the semantics the multi-host test
 * harness measured live (run-vba-tests.ps1), adjusted for F5's two
 * differences: the host opens VISIBLY so the user watches the macro run and
 * sees any error dialog Word raises for a runtime error, and a zero-argument
 * Run needs none of the harness's argument marshaling ([ref] for Word's ByRef
 * varargs; PowerPoint still runs through reflection, the measured-safe route
 * around its ParamArray binding).
 */
export function buildWordMacroLaunchScript(filePath: string, macroName: string, attachToRunning = true): string {
    return [
        '$ErrorActionPreference = "Stop"',
        COM_RETRY_HELPER,
        unsavedWorkFunction('word'),
        'try {',
        `$macroName = ${psSingleQuoted(macroName)}`,
        `$targetPath = ${psSingleQuoted(filePath)}`,
        `$targetName = ${psSingleQuoted(path.win32.basename(filePath))}`,
        '$app = $null',
        '$doc = $null',
        `$attachToRunning = ${attachToRunning ? '$true' : '$false'}`,
        'if ($attachToRunning) { try { $app = [Runtime.InteropServices.Marshal]::GetActiveObject("Word.Application") } catch { } }',
        'if (-not $app) { $app = New-Object -ComObject Word.Application }',
        '$app.Visible = $true',
        'foreach ($d in @($app.Documents)) { if (($d.FullName -ieq $targetPath) -or ($d.Name -ieq $targetName)) { $doc = $d; break } }',
        'if ($doc) {',
        '  if (-not $doc.ReadOnly) {',
        '    throw "REOPEN_BLOCKED|The document is already open for editing in Word. Close it in Word, then press F5 again so XLIDE can reopen the saved document before running the macro."',
        '  }',
        '  if (Test-XlideUnsavedWork $doc) {',
        '    throw "REOPEN_BLOCKED|The document is open read-only in Word with changes that were never saved, and reopening it to run the macro would lose them. Save them under another name or close it without saving, then press F5 again."',
        '  }',
        '  try { Invoke-XlideCom { $doc.Close(0) }; $doc = $null } catch { throw ("REOPEN_FAILED|XLIDE could not close the existing read-only document before running the macro: " + $_.Exception.Message) }',
        '}',
        // Documents.Open(FileName, ConfirmConversions, ReadOnly, AddToRecentFiles)
        'try { $doc = Invoke-XlideCom { $app.Documents.Open($targetPath, $false, $true, $false) } } catch { throw ("REOPEN_FAILED|XLIDE could not open the document. If it is open outside XLIDE, close it in Word and try again: " + $_.Exception.Message) }',
        'try { $doc.Activate() } catch { }',
        // Word resolves Module.Proc and rejects document-qualified names.
        'try { Invoke-XlideCom { $app.Run($macroName) } } catch { throw ("RUN_FAILED|XLIDE could not run the macro: " + $_.Exception.Message) }',
        '[Console]::Out.WriteLine("XLIDE_MACRO_OK")',
        '} catch {',
        '  [Console]::Error.WriteLine("XLIDE_MACRO_ERROR|" + ($_.Exception.Message -replace "[\\r\\n]+", " "))',
        '  exit 1',
        '}',
    ].join('\n');
}

/**
 * Run a procedure in an Access database.
 *
 * Access differs from Word and PowerPoint in three ways that all show up here.
 * It holds one database at a time, so opening ours closes whatever else that
 * instance had, and opening the one already open is an error rather than a
 * no-op: the script asks what is open first. There is no read-only reopen to
 * do, because a database is not a document Access holds a private copy of.
 * And `Application.Run` takes the bare procedure name - measured on Access
 * 16.0, where `Module1.Main` is refused with "cannot find the procedure" and
 * `Main` returns its value.
 *
 * Access recompiles a database XLIDE has written on the next open, so a
 * compile error anywhere in the project surfaces here rather than at the call.
 */
export function buildAccessMacroLaunchScript(filePath: string, procedureName: string, attachToRunning = true): string {
	return [
		'$ErrorActionPreference = "Stop"',
		COM_RETRY_HELPER,
		'try {',
		`$macroName = ${psSingleQuoted(procedureName)}`,
		...accessDatabaseLines(filePath, attachToRunning),
		// Access resolves a bare procedure name and refuses a qualified one.
		// F5 targets a Function as readily as a Sub here, so its return value
		// is swallowed rather than printed over the script's own output.
		'try { Invoke-XlideCom { $app.Run($macroName) } | Out-Null } catch { throw ("RUN_FAILED|XLIDE could not run the procedure: " + $_.Exception.Message) }',
		'[Console]::Out.WriteLine("XLIDE_MACRO_OK")',
		'} catch {',
		'  [Console]::Error.WriteLine("XLIDE_MACRO_ERROR|" + ($_.Exception.Message -replace "[\\r\\n]+", " "))',
		'  exit 1',
		'}',
	].join('\n');
}

/**
 * Two Access rules, both measured on Access 16.0. An instance automation
 * started quits the moment the script lets go of it, unless something visible
 * is open: the database vanished as the script ended, where Excel, Word and
 * PowerPoint stay. Handing the instance to the user (UserControl) keeps it,
 * as theirs to close. And an instance the user controls - one they started,
 * or one kept this way - refuses any write to Visible, even to True, with
 * "invalid reference to the property Visible", so that write is best-effort.
 */
const ACCESS_SHOW_AND_KEEP_LINES: readonly string[] = [
	'try { $app.Visible = $true } catch { }',
	'try { $app.UserControl = $true } catch { }',
];

/**
 * Leaves the running (or a new) Access in `$app` with the target database
 * open. A database the instance already holds is closed first, which is how
 * Access works - but never one with an object edited and not saved.
 */
function accessDatabaseLines(filePath: string, attachToRunning: boolean): string[] {
	return [
		unsavedWorkFunction('access'),
		`$targetPath = ${psSingleQuoted(filePath)}`,
		'$app = $null',
		`$attachToRunning = ${attachToRunning ? '$true' : '$false'}`,
		'if ($attachToRunning) { try { $app = [Runtime.InteropServices.Marshal]::GetActiveObject("Access.Application") } catch { } }',
		'if (-not $app) { $app = New-Object -ComObject Access.Application }',
		...ACCESS_SHOW_AND_KEEP_LINES,
		'$open = ""',
		'try { $open = $app.CurrentProject.FullName } catch { }',
		'if ($open -ine $targetPath) {',
		// Access opens one database at a time; the one already there goes first.
		// Not REOPEN_BLOCKED: that asks the caller to close THIS database and
		// retry, which cannot help when the one in the way is another.
		'  if ($open -and (Test-XlideUnsavedWork $app)) { throw "REOPEN_FAILED|Access has another database open with changes that were never saved, and XLIDE would have to close it to open this one. Save or close it in Access, then press F5 again." }',
		'  if ($open) { try { Invoke-XlideCom { $app.CloseCurrentDatabase() } } catch { throw ("REOPEN_FAILED|XLIDE could not close the database Access already had open: " + $_.Exception.Message) } }',
		'  try { Invoke-XlideCom { $app.OpenCurrentDatabase($targetPath) } } catch { throw ("REOPEN_FAILED|XLIDE could not open the database. If another program holds it open, close it and try again: " + $_.Exception.Message) }',
		'}',
	];
}

/**
 * F5 on an Access form or report. Excel, Word and PowerPoint show a UserForm
 * through a launcher macro because nothing else can; an Access design is a
 * database object, and Access opens one by name, so nothing goes into the
 * database. A report opens in print preview, the view that prints nothing.
 */
export function buildAccessShowDesignScript(
	filePath: string,
	design: { kind: 'form' | 'report'; name: string },
	attachToRunning = true,
): string {
	const noun = design.kind;
	return [
		'$ErrorActionPreference = "Stop"',
		COM_RETRY_HELPER,
		'try {',
		`$designName = ${psSingleQuoted(design.name)}`,
		...accessDatabaseLines(filePath, attachToRunning),
		design.kind === 'form'
			? `try { Invoke-XlideCom { $app.DoCmd.OpenForm($designName) } } catch { throw ("RUN_FAILED|XLIDE could not open the ${noun}: " + $_.Exception.Message) }`
			// acViewPreview
			: `try { Invoke-XlideCom { $app.DoCmd.OpenReport($designName, 2) } } catch { throw ("RUN_FAILED|XLIDE could not open the ${noun}: " + $_.Exception.Message) }`,
		...foregroundLines('$app.hWndAccessApp()'),
		'[Console]::Out.WriteLine("XLIDE_MACRO_OK")',
		'} catch {',
		'  [Console]::Error.WriteLine("XLIDE_MACRO_ERROR|" + ($_.Exception.Message -replace "[\\r\\n]+", " "))',
		'  exit 1',
		'}',
	].join('\n');
}

export function buildPowerPointMacroLaunchScript(filePath: string, macroName: string): string {
    return [
        '$ErrorActionPreference = "Stop"',
        COM_RETRY_HELPER,
        unsavedWorkFunction('powerpoint'),
        'try {',
        `$macroName = ${psSingleQuoted(macroName)}`,
        `$targetPath = ${psSingleQuoted(filePath)}`,
        `$targetName = ${psSingleQuoted(path.win32.basename(filePath))}`,
        // PowerPoint is single-instance: New-Object hands back the user's
        // running application when one exists.
        '$app = New-Object -ComObject PowerPoint.Application',
        '$pres = $null',
        'foreach ($p in @($app.Presentations)) { if (($p.FullName -ieq $targetPath) -or ($p.Name -ieq $targetName)) { $pres = $p; break } }',
        'if ($pres) {',
        '  if ($pres.ReadOnly -eq 0) {',
        '    throw "REOPEN_BLOCKED|The presentation is already open for editing in PowerPoint. Close it in PowerPoint, then press F5 again so XLIDE can reopen the saved presentation before running the macro."',
        '  }',
        '  if (Test-XlideUnsavedWork $pres) {',
        '    throw "REOPEN_BLOCKED|The presentation is open read-only in PowerPoint with changes that were never saved, and reopening it to run the macro would lose them. Save them under another name or close it without saving, then press F5 again."',
        '  }',
        '  try { Invoke-XlideCom { $pres.Close() }; $pres = $null } catch { throw ("REOPEN_FAILED|XLIDE could not close the existing read-only presentation before running the macro: " + $_.Exception.Message) }',
        '}',
        // Presentations.Open(FileName, ReadOnly:=msoTrue, Untitled:=msoFalse,
        // WithWindow:=msoTrue) - visible, unlike the windowless test host.
        'try { $pres = Invoke-XlideCom { $app.Presentations.Open($targetPath, -1, 0, -1) } } catch { throw ("REOPEN_FAILED|XLIDE could not open the presentation. If it is open outside XLIDE, close it in PowerPoint and try again: " + $_.Exception.Message) }',
        'try { $pres.Windows.Item(1).Activate() } catch { }',
        // PowerPoint takes the presentation-qualified File.pptm!Module.Proc form.
        '$macroRef = $pres.Name + "!" + $macroName',
        'try { Invoke-XlideCom { [void]$app.GetType().InvokeMember("Run", [Reflection.BindingFlags]::InvokeMethod, $null, $app, @($macroRef)) } } catch { throw ("RUN_FAILED|XLIDE could not run the macro: " + $_.Exception.Message) }',
        '[Console]::Out.WriteLine("XLIDE_MACRO_OK")',
        '} catch {',
        '  [Console]::Error.WriteLine("XLIDE_MACRO_ERROR|" + ($_.Exception.Message -replace "[\\r\\n]+", " "))',
        '  exit 1',
        '}',
    ].join('\n');
}

export interface HostOpenScriptOptions {
    host: OfficeHostApp;
    filePath: string;
    attachToRunning: boolean;
    /** Access has no read-only open, so a database ignores this. */
    readOnly: boolean;
    /**
     * Leave the application where it is instead of bringing it to the front:
     * for putting back a copy a save closed, while the user works elsewhere.
     */
    background?: boolean;
    /** Where the user was in the copy being put back (officeViewPlace.ts). */
    place?: OfficeViewPlace;
}

/**
 * Opens a file in the application that owns it, reusing a copy that is already
 * open, and brings it to the front. The open call and the window handle of
 * each application were measured live on Office 16.0.
 */
export function buildHostOpenScript(options: HostOpenScriptOptions): string {
    const { host, filePath, attachToRunning, readOnly, background, place } = options;
    if (host === 'excel') {
        return buildExcelLaunchScript({ filePath, attachToRunning, mode: { kind: 'open', readOnly, background, place } });
    }
    const target = [
        '$ErrorActionPreference = "Stop"',
        COM_RETRY_HELPER,
        unsavedWorkFunction(host),
        '$openState = "opened"',
        `$targetPath = ${psSingleQuoted(filePath)}`,
        `$targetName = ${psSingleQuoted(path.win32.basename(filePath))}`,
        '$app = $null',
        '$file = $null',
        `$attachToRunning = ${attachToRunning ? '$true' : '$false'}`,
    ];
    if (host === 'word') {
        return [
            ...target,
            'if ($attachToRunning) { try { $app = [Runtime.InteropServices.Marshal]::GetActiveObject("Word.Application") } catch { } }',
            'if (-not $app) { $app = New-Object -ComObject Word.Application }',
            '$app.Visible = $true',
            'foreach ($d in @($app.Documents)) { if (($d.FullName -ieq $targetPath) -or ($d.Name -ieq $targetName)) { $file = $d; break } }',
            // wdDoNotSaveChanges - reached only when there is nothing unsaved.
            ...(readOnly ? readOnlyReuseLines('$file', '[bool]$file.ReadOnly', 'Invoke-XlideCom { $file.Close(0) }') : []),
            // Documents.Open(FileName, ConfirmConversions, ReadOnly, AddToRecentFiles)
            `if (-not $file) { $file = Invoke-XlideCom { $app.Documents.Open($targetPath, $false, ${readOnly ? '$true' : '$false'}, $false) } }`,
            ...placeRestoreLines(host, place),
            ...(background ? [] : foregroundLines('$app.ActiveWindow.Hwnd', '$file.Activate()')),
            OPEN_SENTINEL_LINE,
        ].join('\n');
    }
    if (host === 'powerpoint') {
        return [
            ...target,
            // PowerPoint is single-instance: New-Object hands back the user's
            // running application when one exists, so there is nothing to attach.
            '$app = New-Object -ComObject PowerPoint.Application',
            'foreach ($p in @($app.Presentations)) { if (($p.FullName -ieq $targetPath) -or ($p.Name -ieq $targetName)) { $file = $p; break } }',
            // MsoTriState: 0 is msoFalse, open for editing.
            ...(readOnly ? readOnlyReuseLines('$file', '$file.ReadOnly -ne 0', 'Invoke-XlideCom { $file.Close() }') : []),
            // Presentations.Open(FileName, ReadOnly, Untitled:=msoFalse, WithWindow:=msoTrue)
            `if (-not $file) { $file = Invoke-XlideCom { $app.Presentations.Open($targetPath, ${readOnly ? '-1' : '0'}, 0, -1) } }`,
            ...placeRestoreLines(host, place),
            ...(background ? [] : foregroundLines('$app.HWND', '$file.Windows.Item(1).Activate()')),
            OPEN_SENTINEL_LINE,
        ].join('\n');
    }
    return [
        ...target,
        'if ($attachToRunning) { try { $app = [Runtime.InteropServices.Marshal]::GetActiveObject("Access.Application") } catch { } }',
        '$open = ""',
        'if ($app) { try { $open = $app.CurrentProject.FullName } catch { } }',
        // One instance holds one database. A running Access that has another
        // one open keeps it, and ours gets its own instance, the way opening
        // it from Explorer would.
        'if ($app -and $open -and ($open -ine $targetPath)) { $app = $null; $open = "" }',
        'if (-not $app) { $app = New-Object -ComObject Access.Application }',
        ...ACCESS_SHOW_AND_KEEP_LINES,
        'if ($open -ine $targetPath) { Invoke-XlideCom { $app.OpenCurrentDatabase($targetPath) } }',
        ...(background ? [] : foregroundLines('$app.hWndAccessApp()')),
        OPEN_SENTINEL_LINE,
    ].join('\n');
}

// Backstop so an open/macro-run COM call cannot hang (and leak its powershell +
// COM reference) forever - the common cause is a modal dialog (e.g. a MsgBox in
// the macro) waiting for the user. Generous so a legitimately long-running
// macro is not cut short.
const HOST_SCRIPT_TIMEOUT_MS = 300_000;

async function runHostScript(
    script: string,
    tag: string,
    log: (message: string) => void,
): Promise<{ code: number | null; stdoutLines: string[]; stderrLines: string[]; spawnError?: Error; timedOut: boolean }> {
    const run = runPowerShell({
        script,
        timeoutMs: HOST_SCRIPT_TIMEOUT_MS,
        onSpawn: (pid) => log(`[${tag}] Spawned powershell.exe (pid=${pid ?? 'unknown'})`),
        onStdoutLine: (line) => log(`[${tag} stdout] ${line}`),
        onStderrLine: (line) => log(`[${tag} stderr] ${line}`),
    });
    const result = await run.result;
    if (!result.spawnError) {
        log(`[${tag}] powershell exited with code=${result.code} signal=${result.signal ?? 'none'}${result.timedOut ? ' (timed out)' : ''}`);
    }
    return result;
}

/** The Office application a launch targets; a VB6 project has none. */
function requireOfficeHost(filePath: string): OfficeHostApp {
    const host = officeHostForPath(filePath);
    if (!host) {
        throw new Error(`${path.basename(filePath)} is not an Office file, so no Office application opens it.`);
    }
    return host;
}

/**
 * What an open left behind that the caller should tell the user about, for
 * an open asked to be read-only. Everything absent means it went as asked.
 */
export interface HostOpenOutcome {
    /**
     * The file is still open for editing: a copy already open that way held
     * unsaved work, so it was left alone ('unsaved'), or it could not be
     * closed ('couldNotClose'). Either way the file stays locked.
     */
    keptEditing?: 'unsaved' | 'couldNotClose';
    /** Excel: the read-only copy opened, but something else still has the file open for editing. */
    lockedElsewhere?: boolean;
    /** Access has no read-only open, so the database opened for editing. */
    noReadOnlyOpen?: boolean;
}

/** Reads the open script's sentinel. */
export function hostOpenOutcome(host: OfficeHostApp, readOnly: boolean, stdoutLines: readonly string[]): HostOpenOutcome {
    if (!readOnly) {
        return {};
    }
    if (host === 'access') {
        return { noReadOnlyOpen: true };
    }
    const state = stdoutLines.find((line) => line.startsWith(OPEN_SENTINEL))?.slice(OPEN_SENTINEL.length).trim();
    switch (state) {
        case 'keptUnsaved':
            return { keptEditing: 'unsaved' };
        case 'keptEditing':
            return { keptEditing: 'couldNotClose' };
        case 'lockedElsewhere':
            return { lockedElsewhere: true };
        default:
            return {};
    }
}

/**
 * Opens (or re-foregrounds) the file in the application that owns it. Rejects
 * only on a spawn failure.
 */
export async function openFileInHost(
    filePath: string,
    options: Pick<HostOpenScriptOptions, 'attachToRunning' | 'readOnly' | 'background' | 'place'>,
    log: (message: string) => void,
): Promise<HostOpenOutcome> {
    const host = requireOfficeHost(filePath);
    const appName = OFFICE_HOST_APPS[host].noun;
    const script = buildHostOpenScript({ host, filePath, ...options });
    log(`[openInHost] Opening in ${appName}. Script:\n${script}`);
    const result = await runHostScript(script, 'openInHost', log);
    if (result.spawnError) {
        log(`[openInHost] Error: ${result.spawnError.message}`);
        throw result.spawnError;
    }
    if (result.timedOut) {
        log(`[openInHost] Timed out and the launch process was killed; ${appName} may be showing a dialog.`);
    }
    return hostOpenOutcome(host, options.readOnly, result.stdoutLines);
}

/**
 * How each application's `Run` names a procedure: Access takes the bare name,
 * the others `Module.Proc` (PowerPoint's script adds the presentation itself).
 */
export function hostMacroReference(host: OfficeHostApp, moduleName: string, procedureName: string): string {
    return host === 'access' ? procedureName : `${moduleName}.${procedureName}`;
}

/**
 * Reopens the file in its own visible application - read-only where the
 * application has such a thing - and runs the procedure. Rejects with a
 * HostMacroError carrying the same codes for every host.
 */
export async function runHostMacro(
    filePath: string,
    macro: { moduleName: string; procedureName: string },
    options: { attachToRunning: boolean },
    log: (message: string) => void,
): Promise<void> {
    const host = requireOfficeHost(filePath);
    const appName = OFFICE_HOST_APPS[host].noun;
    const macroName = hostMacroReference(host, macro.moduleName, macro.procedureName);
    const script = host === 'excel'
        ? buildExcelLaunchScript({
            filePath,
            attachToRunning: options.attachToRunning,
            mode: { kind: 'macroReadOnly', macroName },
        })
        : host === 'word'
            ? buildWordMacroLaunchScript(filePath, macroName, options.attachToRunning)
            : host === 'access'
                ? buildAccessMacroLaunchScript(filePath, macroName, options.attachToRunning)
                : buildPowerPointMacroLaunchScript(filePath, macroName);
    log(`[runMacro] Running in ${appName}: ${macroName}`);
    log(`[runMacro] Script:\n${script}`);
    const result = await runHostScript(script, 'runMacro', log);
    throwUnlessMacroSucceeded(result, appName, log);
}

/**
 * Opens an Access form or report in its database, in the visible application.
 * Rejects with a HostMacroError, as a macro run does.
 */
export async function showAccessDesign(
    filePath: string,
    design: { kind: 'form' | 'report'; name: string },
    options: { attachToRunning: boolean },
    log: (message: string) => void,
): Promise<void> {
    const script = buildAccessShowDesignScript(filePath, design, options.attachToRunning);
    log(`[runForm] Opening ${design.kind} in Access: ${design.name}`);
    log(`[runForm] Script:\n${script}`);
    const result = await runHostScript(script, 'runForm', log);
    throwUnlessMacroSucceeded(result, 'Access', log);
}

/**
 * Turns a finished macro run into the error it stands for, or returns when it
 * succeeded. A spawn failure means PowerShell never ran, so XLIDE never
 * reopened the file: it is UNKNOWN rather than RUN_FAILED, so the F5 handler
 * does not mark the file as XLIDE-opened (a later closeTracked save could
 * then close it out from under the user).
 */
function throwUnlessMacroSucceeded(
    result: Awaited<ReturnType<typeof runHostScript>>,
    appName: string,
    log: (message: string) => void,
): void {
    if (result.spawnError) {
        log(`[runMacro] Error: ${result.spawnError.message}`);
        throw new HostMacroError(result.spawnError.message, 'UNKNOWN');
    }
    if (result.timedOut) {
        throw new HostMacroError(
            `${appName} did not respond within the time limit. A dialog may be open in ${appName} `
            + '(for example a MsgBox or an error dialog from the macro); close it and try again, '
            + 'or the macro may be running too long.',
            'RUN_FAILED',
        );
    }
    if (result.code === 0) {
        return;
    }
    const sentinel = result.stderrLines.find((line) => line.includes(MACRO_ERROR_SENTINEL));
    const raw = sentinel
        ? sentinel.slice(sentinel.indexOf(MACRO_ERROR_SENTINEL) + MACRO_ERROR_SENTINEL.length)
        : result.stderrLines.join('\n') || `PowerShell exited with code ${result.code}`;
    throw hostMacroErrorFromRaw(raw);
}

function hostMacroErrorFromRaw(raw: string): HostMacroError {
    const pipe = raw.indexOf('|');
    const code = pipe >= 0 ? raw.slice(0, pipe) : '';
    if (code === 'REOPEN_BLOCKED' || code === 'REOPEN_FAILED' || code === 'RUN_FAILED') {
        return new HostMacroError(raw.slice(pipe + 1), code);
    }
    return new HostMacroError(raw, 'UNKNOWN');
}
