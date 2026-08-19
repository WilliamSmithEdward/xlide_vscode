import { describe, expect, it } from 'vitest';
import {
    excelComAvailabilityFromProbe,
    excelComProbePowerShellScript,
} from '../src/excelComAvailability';

describe('Excel COM availability', () => {
    it('checks COM registration without creating an Excel instance', () => {
        const script = excelComProbePowerShellScript();

        expect(script).toContain('GetTypeFromProgID("Excel.Application")');
        expect(script).not.toContain('New-Object -ComObject Excel.Application');
        expect(script).not.toContain('GetActiveObject');
    });

    it('reports ready when the Excel COM ProgID is registered', () => {
        expect(excelComAvailabilityFromProbe('win32', 0, 'XLIDE_EXCEL_COM_OK', '')).toEqual({
            state: 'installed',
            title: 'Excel COM Ready',
            description: 'Microsoft Excel is registered for COM automation on this machine.',
            canRun: true,
        });
    });

    it('reports missing when the Excel COM ProgID is absent', () => {
        expect(excelComAvailabilityFromProbe('win32', 2, 'XLIDE_EXCEL_COM_MISSING', '')).toEqual({
            state: 'missing',
            title: 'Excel COM Not Found',
            description: 'Install Microsoft Excel before running VBA tests through XLIDE.',
            canRun: false,
        });
    });

    it('blocks non-Windows test execution before probing Excel', () => {
        expect(excelComAvailabilityFromProbe('linux', null, '', '')).toEqual({
            state: 'blocked',
            title: 'Excel COM Unavailable',
            description: 'VBA tests require Microsoft Excel COM automation on Windows.',
            canRun: false,
        });
    });
});

describe('the probe follows the file host (issue #24)', () => {
    it('probes the ProgID of the requested application', () => {
        expect(excelComProbePowerShellScript('word')).toContain('Word.Application');
        expect(excelComProbePowerShellScript('powerpoint')).toContain('PowerPoint.Application');
        expect(excelComProbePowerShellScript()).toContain('Excel.Application');
    });

    it('names the requested application in every outcome', () => {
        const missing = excelComAvailabilityFromProbe('win32', 2, 'XLIDE_EXCEL_COM_MISSING', '', 'word');
        expect(missing.title).toBe('Word COM Not Found');
        expect(missing.description).toContain('Microsoft Word');
        const ready = excelComAvailabilityFromProbe('win32', 0, 'XLIDE_EXCEL_COM_OK', '', 'powerpoint');
        expect(ready.title).toBe('PowerPoint COM Ready');
        expect(ready.canRun).toBe(true);
    });
});
