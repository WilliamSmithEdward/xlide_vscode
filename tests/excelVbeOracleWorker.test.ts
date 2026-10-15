import * as fs from 'fs';
import * as path from 'path';
import { describe, expect, it } from 'vitest';

describe('Excel VBE oracle worker', () => {
    it('runs Excel visible so the VBE error dialog is detectable', () => {
        // The Win32 dialog watcher only sees VISIBLE top-level windows. When the
        // host Excel is hidden, the VBE Compile-error / Run-time-error modal is
        // not surfaced as a detectable visible window, so every rejection is
        // silently misrecorded as "accepted". The worker therefore keeps Excel
        // visible (and shows the VBE main window so the Compile command and its
        // error dialog appear).
        const worker = fs.readFileSync(
            path.join(__dirname, '..', 'syntax_corpus', 'oracle', 'excel_vbe_oracle_worker.ps1'),
            'utf8',
        );

        expect(worker).toContain('$excel.Visible = $true');
        expect(worker).not.toContain('$excel.Visible = $false');
        expect(worker).toContain('$excel.VBE.MainWindow.Visible = $true');
    });

    it('checks that Excel is alive before accepting without a dialog', () => {
        // A crashed Excel shows no dialog either, so "no dialog" alone recorded
        // a crash as accepted. Every accept that is not backed by a dialog must
        // first confirm that the worker's Excel is still running.
        const worker = fs.readFileSync(
            path.join(__dirname, '..', 'syntax_corpus', 'oracle', 'excel_vbe_oracle_worker.ps1'),
            'utf8',
        );
        const lines = worker.split(/\r?\n/).map((line) => line.trim());
        const accepts = lines
            .map((line, index) => ({ line, index }))
            .filter(({ line }) => line === '$result.outcome = "accepted"');
        const dialogAccepts = accepts.filter(({ index }) => lines[index + 1] === '$result.stage = "vbe_dialog_after_compile"');
        const plainAccepts = accepts.filter((accept) => !dialogAccepts.includes(accept));

        expect(dialogAccepts).toHaveLength(1);
        expect(plainAccepts.length).toBeGreaterThan(0);
        for (const { index } of plainAccepts) {
            expect(lines[index - 1]).toBe('Assert-ExcelAlive $excel $excelPid');
        }

        const assertStart = worker.indexOf('function Assert-ExcelAlive');
        const assertBody = worker.slice(assertStart, worker.indexOf('\nfunction ', assertStart));
        expect(assertBody).toContain('Get-Process -Id $ExcelPid');
        expect(assertBody).toContain('$null = $Excel.Hwnd');
        expect(assertBody.match(/throw "/g)).toHaveLength(2);
    });
});
