import { describe, expect, it } from 'vitest';
import {
    buildAccessMacroLaunchScript,
    buildAccessShowDesignScript,
    buildExcelLaunchScript,
    buildHostOpenScript,
    buildPowerPointMacroLaunchScript,
    buildWordMacroLaunchScript,
    hostMacroReference,
} from '../src/officeHostLauncher';

describe('Excel launcher script', () => {
    const openScript = buildExcelLaunchScript({
        filePath: 'C:\\work\\Book.xlsm',
        attachToRunning: true,
        mode: { kind: 'open', readOnly: false },
    });
    const macroScript = buildExcelLaunchScript({
        filePath: 'C:\\work\\Book.xlsm',
        attachToRunning: true,
        mode: { kind: 'macroReadOnly', macroName: 'Module1.Main' },
    });

    it('shares the attach and foreground fragments between open and macro modes', () => {
        const sharedFragments = [
            '$excel = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")',
            '$excel = New-Object -ComObject Excel.Application',
            'if (($wb.FullName -ieq $targetPath) -or ($wb.Name -ieq $targetName)) { $workbook = $wb; break }',
            '$workbook.Activate()',
            '[XlideHelper.XlideWin32]::SetForegroundWindow([IntPtr]$excel.Hwnd)',
        ];
        for (const fragment of sharedFragments) {
            expect(openScript).toContain(fragment);
            expect(macroScript).toContain(fragment);
        }
    });

    it('opens with the requested read-only flag and no macro sentinels', () => {
        expect(openScript).toContain('$excel.Workbooks.Open($targetPath, 0, $false)');
        expect(openScript).not.toContain('XLIDE_MACRO_ERROR|');

        const readOnlyScript = buildExcelLaunchScript({
            filePath: 'C:\\work\\Book.xlsm',
            attachToRunning: false,
            mode: { kind: 'open', readOnly: true },
        });
        expect(readOnlyScript).toContain('$excel.Workbooks.Open($targetPath, 0, $true)');
        expect(readOnlyScript).toContain('$attachToRunning = $false');
    });

    it('runs the macro behind a read-only reopen with the failure sentinels', () => {
        expect(macroScript).toContain("$macroName = 'Module1.Main'");
        expect(macroScript).toContain('REOPEN_BLOCKED|');
        expect(macroScript).toContain('REOPEN_FAILED|');
        expect(macroScript).toContain('RUN_FAILED|');
        expect(macroScript).toContain('$excel.Run($macroRef)');
        expect(macroScript).toContain('XLIDE_MACRO_ERROR|');
    });

    it('retries Excel-busy COM rejections (RPC_E_CALL_REJECTED) around the macro run and open', () => {
        // The retry helper is present and gates on the busy HResults / messages.
        expect(macroScript).toContain('function Invoke-XlideCom');
        expect(macroScript).toContain('-2147418111');
        expect(macroScript).toContain('rejected by callee');
        // The Run, reopen, and close go through the retry wrapper.
        expect(macroScript).toContain('Invoke-XlideCom { $excel.Run($macroRef) }');
        expect(macroScript).toContain('Invoke-XlideCom { $excel.Workbooks.Open($targetPath, 0, $true) }');
        expect(macroScript).toContain('Invoke-XlideCom { $workbook.Close($false) }');
        // Foreground Activate is best-effort, never failing the run on a busy Excel.
        expect(macroScript).toContain('try { $workbook.Activate() } catch { }');
        // Plain open also retries its Open call.
        expect(openScript).toContain('Invoke-XlideCom { $excel.Workbooks.Open($targetPath, 0, $false) }');
    });

    it('escapes single quotes in interpolated values', () => {
        const script = buildExcelLaunchScript({
            filePath: "C:\\work\\Bob's Book.xlsm",
            attachToRunning: false,
            mode: { kind: 'open', readOnly: false },
        });
        expect(script).toContain("$targetPath = 'C:\\work\\Bob''s Book.xlsm'");
        expect(script).toContain("$targetName = 'Bob''s Book.xlsm'");
    });
});

describe('Word and PowerPoint F5 macro launcher scripts', () => {
    // Live-verified 2026-08-19: both scripts ran their macro in the real
    // application (marker file written, XLIDE_MACRO_OK), rejected a missing
    // macro as RUN_FAILED, and exercised the stale read-only close-and-reopen
    // path on their second run.
    const wordScript = buildWordMacroLaunchScript('C:\\work\\Report.docm', 'Module1.Main');
    const pptScript = buildPowerPointMacroLaunchScript('C:\\work\\Deck.pptm', 'Module1.Main');

    it('drives Word with the measured open, visible window, and plain Module.Proc run', () => {
        expect(wordScript).toContain('GetActiveObject("Word.Application")');
        expect(wordScript).toContain('New-Object -ComObject Word.Application');
        expect(wordScript).toContain('$app.Visible = $true');
        // Documents.Open(FileName, ConfirmConversions, ReadOnly, AddToRecentFiles)
        expect(wordScript).toContain('$app.Documents.Open($targetPath, $false, $true, $false)');
        // Zero-argument run: no [ref] marshaling, no document qualifier.
        expect(wordScript).toContain('$app.Run($macroName)');
        expect(wordScript).not.toContain('[ref]');
        expect(wordScript).toContain('REOPEN_BLOCKED|The document is already open for editing in Word.');
        expect(wordScript).toContain('XLIDE_MACRO_OK');
        expect(wordScript).toContain('XLIDE_MACRO_ERROR|');
    });

    it('drives PowerPoint single-instance with a visible window and reflection Run', () => {
        // Single-instance host: New-Object attaches to a running PowerPoint.
        expect(pptScript).toContain('New-Object -ComObject PowerPoint.Application');
        expect(pptScript).not.toContain('GetActiveObject("PowerPoint.Application")');
        // Presentations.Open(FileName, ReadOnly, Untitled, WithWindow=msoTrue).
        expect(pptScript).toContain('$app.Presentations.Open($targetPath, -1, 0, -1)');
        // Presentation-qualified name through reflection InvokeMember.
        expect(pptScript).toContain('$macroRef = $pres.Name + "!" + $macroName');
        expect(pptScript).toContain('InvokeMember("Run", [Reflection.BindingFlags]::InvokeMethod, $null, $app, @($macroRef))');
        expect(pptScript).toContain('REOPEN_BLOCKED|The presentation is already open for editing in PowerPoint.');
        expect(pptScript).toContain('XLIDE_MACRO_OK');
    });

    it('drives Access by the database it holds and the bare procedure name', () => {
        // Live-verified 2026-09-05 on Access 16.0: the script ran Main and
        // printed XLIDE_MACRO_OK on a fresh open, on a second run with the
        // same database still open, and after switching Access from another
        // database; a missing procedure came back RUN_FAILED with exit 1.
        const script = buildAccessMacroLaunchScript('C:\\work\\Orders.accdb', 'Main');
        expect(script).toContain('GetActiveObject("Access.Application")');
        expect(script).toContain('New-Object -ComObject Access.Application');
        expect(script).toContain('$app.Visible = $true');
        // One database at a time: what is open is asked for, and closed only
        // when it is a different one.
        expect(script).toContain('$open = $app.CurrentProject.FullName');
        expect(script).toContain('if ($open -ine $targetPath) {');
        expect(script).toContain('$app.CloseCurrentDatabase()');
        expect(script).toContain('$app.OpenCurrentDatabase($targetPath)');
        // Access refuses Module.Proc and resolves the bare name.
        expect(script).toContain('$app.Run($macroName)');
        expect(script).toContain("$macroName = 'Main'");
        expect(script).not.toContain('[ref]');
        // A database is not reopened read-only; there is no private copy.
        expect(script).not.toContain('REOPEN_BLOCKED');
        // Measured 2026-09-16 on Access 16.0. An Access the user started
        // refuses any write to Visible ("invalid reference to the property
        // Visible"), which failed every F5 against it, so the write is
        // best-effort. And an Access the script started quits as the script
        // ends unless it is handed to the user, which is what keeps the
        // database open afterward the way the other applications keep a file.
        expect(script).toContain('try { $app.Visible = $true } catch { }');
        expect(script).toContain('try { $app.UserControl = $true } catch { }');
        expect(script).toContain('XLIDE_MACRO_OK');
        expect(script).toContain('XLIDE_MACRO_ERROR|');
    });

    it('escapes single quotes in interpolated host values', () => {
        const script = buildWordMacroLaunchScript("C:\\work\\Bob's Report.docm", 'Module1.Main');
        expect(script).toContain("$targetPath = 'C:\\work\\Bob''s Report.docm'");
        expect(script).toContain("$targetName = 'Bob''s Report.docm'");
    });

    it('lets Word and Access start their own instance when attaching is turned off', () => {
        // The attach is the same measured line, now behind the setting Excel
        // already honored. PowerPoint only ever runs one instance.
        expect(wordScript).toContain('$attachToRunning = $true');
        expect(wordScript).toContain('if ($attachToRunning) { try { $app = [Runtime.InteropServices.Marshal]::GetActiveObject("Word.Application") } catch { } }');
        expect(buildWordMacroLaunchScript('C:\\work\\Report.docm', 'Module1.Main', false))
            .toContain('$attachToRunning = $false');
        expect(buildAccessMacroLaunchScript('C:\\work\\Orders.accdb', 'Main', false))
            .toContain('$attachToRunning = $false');
    });

    it('names a procedure the way each application resolves it', () => {
        expect(hostMacroReference('excel', 'Module1', 'Main')).toBe('Module1.Main');
        expect(hostMacroReference('word', 'Module1', 'Main')).toBe('Module1.Main');
        expect(hostMacroReference('powerpoint', 'Module1', 'Main')).toBe('Module1.Main');
        // Access refuses a qualified name.
        expect(hostMacroReference('access', 'Module1', 'Main')).toBe('Main');
    });
});

describe('open-in-application scripts', () => {
    // Live-verified 2026-09-16 on Office 16.0 against scratch copies of the
    // fixtures: each script opened its file in the visible application, for
    // editing and read-only, and the application's own ReadOnly property
    // agreed with what was asked for.
    const open = (filePath: string, readOnly: boolean, attachToRunning = true): string => {
        const host = /\.docm$/.test(filePath) ? 'word' : /\.pptm$/.test(filePath) ? 'powerpoint'
            : /\.accdb$/.test(filePath) ? 'access' : 'excel';
        return buildHostOpenScript({ host, filePath, attachToRunning, readOnly });
    };

    it('opens an Excel workbook with the launcher script Excel always had', () => {
        expect(open('C:\\work\\Book.xlsm', true)).toBe(buildExcelLaunchScript({
            filePath: 'C:\\work\\Book.xlsm',
            attachToRunning: true,
            mode: { kind: 'open', readOnly: true },
        }));
    });

    it('opens a Word document editable or read-only, reusing an open copy', () => {
        const editable = open('C:\\work\\Report.docm', false);
        expect(editable).toContain('GetActiveObject("Word.Application")');
        expect(editable).toContain('New-Object -ComObject Word.Application');
        expect(editable).toContain('foreach ($d in @($app.Documents))');
        // Documents.Open(FileName, ConfirmConversions, ReadOnly, AddToRecentFiles)
        expect(editable).toContain('if (-not $file) { $file = Invoke-XlideCom { $app.Documents.Open($targetPath, $false, $false, $false) } }');
        expect(open('C:\\work\\Report.docm', true)).toContain('$app.Documents.Open($targetPath, $false, $true, $false)');
        expect(editable).toContain('SetForegroundWindow([IntPtr]$app.ActiveWindow.Hwnd)');
    });

    it('opens a PowerPoint presentation through its single instance', () => {
        const editable = open('C:\\work\\Deck.pptm', false);
        expect(editable).toContain('$app = New-Object -ComObject PowerPoint.Application');
        expect(editable).not.toContain('GetActiveObject');
        // Presentations.Open(FileName, ReadOnly, Untitled, WithWindow): MsoTriState.
        expect(editable).toContain('$app.Presentations.Open($targetPath, 0, 0, -1)');
        expect(open('C:\\work\\Deck.pptm', true)).toContain('$app.Presentations.Open($targetPath, -1, 0, -1)');
        expect(editable).toContain('SetForegroundWindow([IntPtr]$app.HWND)');
    });

    it('opens an Access database without taking over an instance that holds another one', () => {
        const script = open('C:\\work\\Orders.accdb', true);
        expect(script).toContain('GetActiveObject("Access.Application")');
        // One instance holds one database: a running Access that has another
        // one open keeps it, and ours gets its own instance.
        expect(script).toContain('if ($app -and $open -and ($open -ine $targetPath)) { $app = $null; $open = "" }');
        expect(script).not.toContain('CloseCurrentDatabase');
        expect(script).toContain('if ($open -ine $targetPath) { Invoke-XlideCom { $app.OpenCurrentDatabase($targetPath) } }');
        // Without this the database would close again as the script ended.
        expect(script).toContain('try { $app.UserControl = $true } catch { }');
        expect(script).toContain('SetForegroundWindow([IntPtr]$app.hWndAccessApp())');
    });

    it('honors the attach setting where the application can run more than one instance', () => {
        expect(open('C:\\work\\Report.docm', false, false)).toContain('$attachToRunning = $false');
        expect(open('C:\\work\\Orders.accdb', false, false)).toContain('$attachToRunning = $false');
    });
});

describe('Access form and report F5 script', () => {
    it('opens a form by name, in the database, with nothing written into it', () => {
        const script = buildAccessShowDesignScript('C:\\work\\Orders.accdb', { kind: 'form', name: "Bob's Orders" });
        expect(script).toContain("$designName = 'Bob''s Orders'");
        expect(script).toContain('$app.OpenCurrentDatabase($targetPath)');
        expect(script).toContain('Invoke-XlideCom { $app.DoCmd.OpenForm($designName) }');
        expect(script).toContain('RUN_FAILED|XLIDE could not open the form: ');
        expect(script).toContain('XLIDE_MACRO_OK');
        expect(script).not.toContain('$app.Run(');
    });

    it('opens a report in print preview, the view that prints nothing', () => {
        const script = buildAccessShowDesignScript('C:\\work\\Orders.accdb', { kind: 'report', name: 'Monthly' });
        // acViewPreview = 2
        expect(script).toContain('Invoke-XlideCom { $app.DoCmd.OpenReport($designName, 2) }');
        expect(script).toContain('RUN_FAILED|XLIDE could not open the report: ');
    });
});
