import { describe, expect, it } from 'vitest';
import {
    buildExcelLaunchScript,
    buildPowerPointMacroLaunchScript,
    buildWordMacroLaunchScript,
} from '../src/excelLauncher';

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

    it('escapes single quotes in interpolated host values', () => {
        const script = buildWordMacroLaunchScript("C:\\work\\Bob's Report.docm", 'Module1.Main');
        expect(script).toContain("$targetPath = 'C:\\work\\Bob''s Report.docm'");
        expect(script).toContain("$targetName = 'Bob''s Report.docm'");
    });
});
