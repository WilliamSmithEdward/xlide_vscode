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
});
