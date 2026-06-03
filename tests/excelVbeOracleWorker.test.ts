import * as fs from 'fs';
import * as path from 'path';
import { describe, expect, it } from 'vitest';

describe('Excel VBE oracle worker', () => {
    it('keeps the Excel application window hidden', () => {
        const worker = fs.readFileSync(
            path.join(__dirname, '..', 'syntax_corpus', 'oracle', 'excel_vbe_oracle_worker.ps1'),
            'utf8',
        );

        expect(worker).toContain('$excel.Visible = $false');
        expect(worker).not.toContain('$excel.Visible = $true');
        expect(worker).toContain('$excel.VBE.MainWindow.Visible = $true');
    });
});
